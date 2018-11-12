"""A module containing a power flow analysis API.

The power flow expects bus and line data to be provided in separate Excel worksheets. Once the analysis parameters and
data are provided, the following algorithm is executed:

    1. load bus data and set flat start voltages
    2. create admittance matrix from line data
    3. until a solution is found:
        3.1 generate power error terms (dP and dQ)
            3.1.1 generate real and reactive power error terms for load buses.
            3.1.2 generate real power error terms for generator buses.
        3.2 generate voltage and phase angle corrections through an iteration of the Newton-Raphson algorithm
            3.2.1 generate the Jacobian matrix (J = [[dP/dtheta, dP/dV], [dQ/dtheta, dQ/dV]])
            3.2.2 solve inv(J) * [[dP], [dQ]] to get a vector of voltage corrections ([[dtheta], [dV]])
        3.3 apply corrections to each bus (except the slack bus)

More formally, let S(actual) = P(actual) + jQ(actual) be a vector of power delivery at each bus, V be a vector of
voltages (and voltage angles) at each bus, and Y be the admittance matrix for the system. Initially, any unknown
voltages in V are set to a reference voltage (usually 1.0 pu, the flat start voltage), and the Newton-Raphson method is
used to reduce the error in bus power to zero, and iterate the system towards a steady-state solution.

    while true:
        P(calculated) = [sum(i = 1 to N) B[k][i] V[k] V[i] sin(theta[k] - theta[i])]
        Q(calculated) = [sum(i = 1 to N) -B[k][i] V[k] V[i] cos(theta[k] - theta[i])]
        P(error) = P(calculated) - P(actual)
        Q(error) = Q(calculated) - Q(actual)

        if P(error) < P(max error) and Q(error) < Q(max error):
            break

        x = [angle(V), abs(V)]
        J = [[dP(error)/dtheta, dP(error)/dV], [dQ(error)/dtheta, dQ(error)/dV]]
        dx = inverse(J) x

        V += dx[0] e^(j dx[1])

The algorithm is considered complete when the magnitude of the error is less than some specified allowable threshold
(max_error_mw and max_error_mvar).
"""

import enum
import namedlist
import numpy
import operator
import power_flow_jacobian


class BusType(enum.Enum):
    """An enumeration of bus types."""
    UNKNOWN = 0
    LOAD = 1  # P and Q consumed are specified.
    GENERATOR = 2  # P delivered and V are specified
    SLACK = 3


# An object containing bus data.
#
# Attributes:
#   type: The bus type.
#   number: The bus number.
#   power: The complex power injection at the bus (pu).
#   voltage: The bus voltage (pu).
Bus = namedlist.namedlist('Bus', ['type', 'number', 'power', 'voltage'])


class PowerFlowException(Exception):
    """An exception class for power flow operations."""
    pass


def execute_power_flow(
    bus_data_ws, line_data_ws, slack_bus_number, start_voltage, power_base_mva, max_error_mw, max_error_mvar):
    """Creates a new power flow object.

    Args:
        bus_data_ws: The openpyxl worksheet containing bus data.
        line_data_ws: The openpyxl worksheet containing line data.
        slack_bus_number: The slack bus number.
        start_voltage: The initial voltages to use at each bus.
        power_base_mva: The power base in MVA.
        max_error_mw: The maximum allowable real power mismatch for a solution to be considered final.
        max_error_mvar: The maximum allowable reactive power mismatch for a solution to be considered final.

    Returns:
        An array of buses and their steady-state solutions.

    Raises:
        PowerFlowException: Indicates that bus data could not be loaded.
    """
    max_active_power_error_pu = max_error_mw / power_base_mva
    max_reactive_power_error_pu = max_error_mvar / power_base_mva

    # TODO(kjiwa): Store line data.
    bus_states = read_bus_states(bus_data_ws, slack_bus_number, start_voltage, power_base_mva)
    admittances = read_admittance_matrix(line_data_ws, len(bus_states))

    while True:
        # Create an estimated state and compute the error.
        bus_estimates = create_bus_estimates(bus_states, admittances)
        bus_errors = calculate_state_error(bus_states, bus_estimates)
        if is_convergent(bus_errors, max_active_power_error_pu, max_reactive_power_error_pu):
            break

        # Compute Jacobian and its inverse.
        jacobian = power_flow_jacobian.create_jacobian_matrix(bus_estimates, admittances)
        invjacobian = numpy.linalg.inv(jacobian)

        # Create the error vector [dP dQ].
        del bus_errors[slack_bus_number]
        sorted_bus_errors = sorted(bus_errors.values(), key=operator.attrgetter('number'))
        pq_error = numpy.transpose([bus.power.real for bus in sorted_bus_errors]
                                   + [bus.power.imag for bus in sorted_bus_errors if bus.type == BusType.LOAD])

        # Compute the corrections. There are (m + n) theta corrections and m magnitude corrections.
        corrections = -numpy.matmul(invjacobian, pq_error)
        theta_corrections = corrections[0:len(sorted_bus_errors)]
        voltage_corrections = corrections[len(sorted_bus_errors):]

        # Apply corrections.
        apply_theta_corrections(bus_states, theta_corrections, sorted_bus_errors)
        apply_voltage_corrections(bus_states, voltage_corrections, sorted_bus_errors)

        # Recompute generator buses and the slack bus power.
        for bus in bus_estimates.values():
            if bus.type == BusType.GENERATOR:
                bus_states[bus.number].power = bus_states[bus.number].power.real + 1j * bus.power.imag
            elif bus.type == BusType.SLACK:
                bus_states[bus.number].power = bus.power

    # Convert power values to MVA.
    for bus in bus_states.values():
        bus.power *= power_base_mva

    return bus_states.values()


def read_bus_states(bus_data_ws, slack_bus_number, start_voltage, power_base_mva):
    """Reads bus data from an input worksheet.

    The input data is expected to have the following columns:

        1. Bus number
        2. Real power consumed (MW)
        3. Reactive power consumed (Mvar)
        4. Real power delivered (MW)
        5. Voltage (pu)

    Returns:
        An array containing bus data.

    Raises:
        PowerFlowException: Indicates that bus data could not be loaded.
    """
    s_total = 0j
    buses = {}
    for i, row in enumerate(bus_data_ws.iter_rows(row_offset=1)):
        bus_num = row[0].value
        p_load = row[1].value or 0
        q_load = row[2].value or 0
        p_gen = row[3].value or 0
        voltage = row[4].value or start_voltage
        voltage += 0j  # force the voltage value to be complex

        if not bus_num:
            break

        # Determine the bus type.
        type = BusType.UNKNOWN
        if bus_num == slack_bus_number:
            type = BusType.SLACK
        elif p_load != 0 or q_load != 0:
            type = BusType.LOAD
        elif p_gen != 0 and numpy.abs(voltage) > 0:
            type = BusType.GENERATOR

        if type == BusType.UNKNOWN:
            raise PowerFlowException('Unable to determine bus type for row {}'.format(i))

        s = (p_gen - (p_load + 1j * q_load)) / power_base_mva
        s_total += s
        buses[bus_num] = Bus(type, bus_num, s, voltage)

    # Balance system power in the slack bus.
    buses[slack_bus_number].power = -s_total
    return buses


def read_admittance_matrix(line_data_ws, num_buses):
    """Reads an admittance matrix from an input worksheet.

    The worksheet is expected to have a header row. Each subsequent row is expected to contain the following
    columns:

        1. Source bus
        2. Destination bus
        3. Total resistance (pu)
        4. Total reactance (pu)
        5. Total susceptance (pu)
        6. Maximum apparent power (MVA)

    Furthermore, it is expected line data is not duplicated and that the source and destination buses are
    different.

    Args:
        num_buses: The number of buses in the system.

    Returns:
        A square matrix containing line-to-line admittances.
    """
    admittances = numpy.zeros((num_buses, num_buses)) * 0j  # multiply by 0j to make values complex

    # Compute matrix entries.
    for row in line_data_ws.iter_rows(row_offset=1):
        src = row[0].value
        dst = row[1].value
        if not src or not dst:
            break

        # Subtract 1 since arrays start at 0, but buses start at 1.
        src -= 1
        dst -= 1

        rtotal = row[2].value
        xtotal = row[3].value
        btotal = row[4].value
        # fmax = row[5].value

        # Series and parallel impedances.
        y_series = 1 / (rtotal + 1j * xtotal)
        y_parallel = 1j * btotal / 2

        # Subtract series admittances from off-diagonal entries.
        admittances[src][dst] -= y_series
        admittances[dst][src] -= y_series

        # Add series and parallel admittances to diagonal entries.
        admittances[src][src] += y_series + y_parallel
        admittances[dst][dst] += y_series + y_parallel

    return admittances


def is_convergent(errors, max_active_power_error_pu, max_reactive_power_error_pu):
    """Checks if the power flow has converged to a solution.

    Args:
        errors: The power error terms.

    Returns:
        True if the power flow has converged to a solution; false otherwise.
    """
    for bus in errors.values():
        if bus.power.real >= max_active_power_error_pu or bus.power.imag >= max_reactive_power_error_pu:
            return False

    return True


def apply_theta_corrections(bus_states, theta_corrections, bus_order):
    """Applies voltage angle corrections to bus states.

    The corrections vector is expected not to have an update for the slack bus. The vector is expected to be sorted
    by bus number.

    Args:
        bus_states: The system buses being updated.
        theta_corrections: A sorted vector of voltage angles.
        bus_order: An list of buses having the same order as the corrections.
    """
    load_gen_bus_numbers = [bus.number for bus in bus_order]
    for bus_number, dtheta in zip(load_gen_bus_numbers, theta_corrections):
        magnitude = numpy.abs(bus_states[bus_number].voltage)
        theta = numpy.angle(bus_states[bus_number].voltage) + dtheta
        bus_states[bus_number].voltage = magnitude * numpy.exp(1j * theta)


def apply_voltage_corrections(bus_states, voltage_corrections, bus_order):
    """Applies voltage magnitude corrections to bus states.

    The corrections vector is expected not to have an update for the slack bus. The vector is expected to be sorted
    by bus number.

    Args:
        bus_states: The system buses being updated.
        voltage_corrections: A sorted vector of voltage magnitudes.
        bus_order: An list of buses having the same order as the corrections.
    """
    load_bus_numbers = [bus.number for bus in bus_order if bus.type == BusType.LOAD]
    for bus_number, dv in zip(load_bus_numbers, voltage_corrections):
        magnitude = numpy.abs(bus_states[bus_number].voltage) + dv
        theta = numpy.angle(bus_states[bus_number].voltage)
        bus_states[bus_number].voltage = magnitude * numpy.exp(1j * theta)


def create_bus_estimates(bus_states, admittances):
    """Estimates the power injection at a particular bus.

    Args:
        bus_states: The current bus state.
        admittances: The system admittance matrix.

    Returns:
        A list of buses with estimates for the power field.
    """
    buses_estimate = {}
    for x, src in bus_states.items():
        s = 0j

        k = src.number - 1
        v_k = numpy.abs(src.voltage)
        theta_k = numpy.angle(src.voltage)

        for y, dst in bus_states.items():
            j = dst.number - 1
            v_j = numpy.abs(dst.voltage)
            theta_j = numpy.angle(dst.voltage)
            b_kj = admittances[k][j].imag

            p = v_k * v_j * b_kj * numpy.sin(theta_k - theta_j)
            q = -v_k * v_j * b_kj * numpy.cos(theta_k - theta_j)
            s += p + 1j * q

        buses_estimate[src.number] = Bus(src.type, src.number, s, src.voltage)

    return buses_estimate


def calculate_state_error(bus_states, bus_estimates):
    """Calculates the error state between the estimated bus power and the actual bus power.

    Args:
        bus_states: The actual bus states.
        bus_estimates: The estimated bus states.

    Returns:
        A list of bus states with error terms in the power field.
    """
    errors = {}
    for bus_number in bus_states.keys():
        actual = bus_states[bus_number]
        estimate = bus_estimates[bus_number]
        diff = actual.power - estimate.power
        errors[bus_number] = Bus(actual.type, bus_number, diff, actual.voltage)

    return errors
