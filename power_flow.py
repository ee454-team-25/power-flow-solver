"""A module containing a power flow analysis API."""

import namedlist
import numpy
import operator

# An object containing bus data.
#
# Attributes:
#   number: The bus number.
#   s: The complex power consumed at the bus (MVA).
#   voltage: The bus voltage (pu).
Bus = namedlist.namedlist('Bus', ['number', 's', 'voltage'])


class PowerFlow():
    """A power flow object."""

    def __init__(
        self, bus_data, line_data, slack_bus_number, start_voltage, power_base_mva, max_mismatch_mw, max_mismatch_mvar):
        """Creates a new power flow object.

        Args:
            bus_data: The openpyxl worksheet containing bus data.
            line_data: The openpyxl worksheet containing line data.
            slack_bus_number: The slack bus number.
            start_voltage: The initial voltages to use at each bus.
            power_base_mva: The power base in MVA.
            max_mismatch_mw: The maximum allowable real power mismatch for a solution to be considered final.
            max_mismatch_mvar: The maximum allowable reactive power mismatch for a solution to be considered final.
        """
        self.bus_data = bus_data
        self.line_data = line_data
        self.slack_bus_number = slack_bus_number
        self.start_voltage = start_voltage
        self.power_base_mva = power_base_mva
        self.max_mismatch_mw = max_mismatch_mw
        self.max_mismatch_mvar = max_mismatch_mvar

    def execute(self):
        """Executes a power flow analysis.

        Returns:
            An array of buses and their steady-state solutions.
        """
        buses = self.read_bus_data()
        admittances = self.read_admittance_matrix(len(buses))

        # TODO(kjiwa): Clean this section up.
        while True:
            # Compute power flow equations.
            s_values = self.create_power_flow_equations(buses, admittances)
            p_values = [i.real for i in s_values]
            q_values = [i.imag for i in s_values]

            # Compute mismatch equations and test for convergence.
            dp_values = [p - buses[i].s.real for i, p in enumerate(p_values) if
                         buses[i].number != self.slack_bus_number]
            dq_values = [q - buses[i].s.imag for i, q in enumerate(q_values) if
                         buses[i].number != self.slack_bus_number]
            if self.is_convergent(dp_values, dq_values):
                break

            # Compute Jacobian and its inverse.
            jacobian = self.create_jacobian_matrix(buses, admittances, p_values, q_values)
            invjacobian = numpy.linalg.inv(jacobian)

            # Compute corrections.
            corrections = -numpy.matmul(invjacobian, numpy.concatenate([dp_values, dq_values]))
            voltage_corrections = corrections[0:len(corrections) // 2]
            theta_corrections = corrections[len(corrections) // 2:]

            # Apply corrections.
            count = 0
            for bus in buses:
                if bus.number == self.slack_bus_number:
                    continue

                voltage = numpy.abs(bus.voltage) + voltage_corrections[count]
                theta = numpy.angle(bus.voltage) + theta_corrections[count]
                bus.voltage = voltage * numpy.exp(1j * theta)
                count += 1

        return buses

    def count_rows(self, ws):
        """Counts the number of non-empty, contiguous rows in the worksheet.

        This method is necessary because openpyxl.worksheet.worksheet.max_row does not always correctly detect empty
        cells.

        Args:
            ws: The worksheet to read.

        Returns:
            The number of non-empty, contiguous rows.
        """
        n = 0
        for row in ws.iter_rows():
            if not row[0].value:
                break

            n += 1

        return n

    def read_bus_data(self):
        """Reads bus data from an input worksheet.

        The input data is expected to have the following columns:

            1. Bus number
            2. Real power consumed (MW)
            3. Reactive power consumed (Mvar)
            4. Real power delivered (MW)
            5. Voltage (pu)

        Returns:
            An array containing bus data.
        """
        s_total = 0j
        buses = []
        for i, row in enumerate(self.bus_data.iter_rows(row_offset=1, max_row=self.count_rows(self.bus_data))):
            bus_num = row[0].value
            p_load = row[1].value or 0
            q_load = row[2].value or 0
            p_gen = row[3].value or 0
            voltage = row[4].value or self.start_voltage
            voltage += 0j  # force the voltage value to be complex

            s = (p_load - p_gen + 1j * q_load) / self.power_base_mva
            s_total += s
            buses.append(Bus(bus_num, s, voltage))

        # Balance system power in the slack bus.
        buses = sorted(buses, key=operator.attrgetter('number'))
        buses[self.slack_bus_number - 1].s = -s_total
        return buses

    def read_admittance_matrix(self, num_buses):
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
        admittances = numpy.zeros((num_buses, num_buses)) * 1j  # multiply by 1j to make values complex

        # Compute matrix entries.
        for row in self.line_data.iter_rows(row_offset=1, max_row=num_buses):
            # Subtract 1 since arrays start at 0, but buses start at 1.
            src = row[0].value - 1
            dst = row[1].value - 1

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

    def create_power_flow_equations(self, buses, admittances):
        """Creates the power flow equations from a set of buses and admittances.

        For each bus, the power flow equations are:

            dPk = sum(i = 1 to N) Vk * Vi * Bki * sin(theta_k - theta_i)
            dQk = sum(i = 1 to N) -Vk * Vi * Bki * cos(theta_k - theta_i)

        N is the number of buses, Vi is the voltage magnitude at bus i, Bki is the line admittance from bus k to bus i,
        and theta_i is the voltage angle at bus i.

        Args:
            buses: The state of the system buses.
            admittances: The system admittances.

        Returns:
            An array of power flow equations in the same order as the input buses.
        """
        s_values = []
        for src in buses:
            s = 0j

            k = src.number - 1
            v_k = numpy.abs(src.voltage)
            theta_k = numpy.angle(src.voltage)

            for dst in buses:
                j = dst.number - 1
                v_j = numpy.abs(dst.voltage)
                theta_j = numpy.angle(dst.voltage)
                b_kj = admittances[k][j].imag

                p = v_k * v_j * b_kj * numpy.sin(theta_k - theta_j)
                q = -v_k * v_j * b_kj * numpy.cos(theta_k - theta_j)
                s += p + 1j * q

            s_values.append(s)

        return s_values

    def create_jacobian_matrix(self, buses, admittances, p_values, q_values):
        """Creates the Jacobian matrix for the state of a power system.

        The Jacobian is a 2n by 2n matrix with the following form:

            J = [[J_11] [J_12]]
                [[J_21] [J_22]]

        Args:
            buses: The state of the system buses.
            admittances: The system admittances.
            p_values: The real power mismatch equations.
            q_values: The reactive power mismatch equations.

        Returns:
            The Jacobian matrix for the power system.
        """
        # Create submatrices.
        j_11 = self.create_j11(buses, admittances, p_values, q_values)
        j_12 = self.create_j12(buses, admittances, p_values, q_values)
        j_21 = self.create_j21(buses, admittances, p_values, q_values)
        j_22 = self.create_j22(buses, admittances, p_values, q_values)

        # Remove slack bus entries.
        j_11 = numpy.delete(j_11, self.slack_bus_number - 1, axis=0)
        j_11 = numpy.delete(j_11, self.slack_bus_number - 1, axis=1)
        j_12 = numpy.delete(j_12, self.slack_bus_number - 1, axis=0)
        j_12 = numpy.delete(j_12, self.slack_bus_number - 1, axis=1)
        j_21 = numpy.delete(j_21, self.slack_bus_number - 1, axis=0)
        j_21 = numpy.delete(j_21, self.slack_bus_number - 1, axis=1)
        j_22 = numpy.delete(j_22, self.slack_bus_number - 1, axis=0)
        j_22 = numpy.delete(j_22, self.slack_bus_number - 1, axis=1)

        # Combine matrices together.
        j_1 = numpy.concatenate([j_11, j_12], axis=1)
        j_2 = numpy.concatenate([j_21, j_22], axis=1)
        j = numpy.concatenate([j_1, j_2])
        return j

    def create_j11(self, buses, admittances, unused_p_values, q_values):
        """Creates the upper-left Jacobian matrix, J_11.

            J_11 = dP_k / dtheta_j

        Args:
            buses: The state of the system buses.
            admittances: The system admittances.
            unused_p_values: The real power mismatch equations.
            q_values: The reactive power mismatch equations.

        Returns:
            The Jacobian matrix J_11.
        """
        j_11 = numpy.zeros((len(buses), len(buses))) * 1j
        for src in buses:
            k = src.number - 1
            v_k = numpy.abs(src.voltage)
            theta_k = numpy.angle(src.voltage)
            q_k = q_values[k]

            for dst in buses:
                j = dst.number - 1
                v_j = numpy.abs(dst.voltage)
                theta_j = numpy.angle(dst.voltage)
                b_kj = admittances[k][j].imag

                if j != k:
                    j_11[k][j] = -v_k * v_j * b_kj * numpy.cos(theta_k - theta_j)
                else:
                    j_11[k][j] = -q_k - v_k ** 2 * b_kj

        return j_11

    def create_j12(self, buses, admittances, p_values, unused_q_values):
        """Creates the upper-left Jacobian matrix, J_12.

            J_12 = dP_k / dV_j

        Args:
            buses: The state of the system buses.
            admittances: The system admittances.
            p_values: The real power mismatch equations.
            unused_q_values: The reactive power mismatch equations.

        Returns:
            The Jacobian matrix J_12.
        """
        j_12 = numpy.zeros((len(buses), len(buses))) * 1j
        for src in buses:
            k = src.number - 1
            v_k = numpy.abs(src.voltage)
            theta_k = numpy.angle(src.voltage)
            p_k = p_values[k]

            for dst in buses:
                j = dst.number - 1
                v_j = numpy.abs(dst.voltage)
                theta_j = numpy.angle(dst.voltage)
                b_kj = admittances[k][j].imag

                if j != k:
                    j_12[k][j] = v_k * b_kj * numpy.sin(theta_k - theta_j)
                else:
                    j_12[k][j] = p_k / v_k

        return j_12

    def create_j21(self, buses, admittances, p_values, unused_q_values):
        """Creates the upper-left Jacobian matrix, J_21.

            J_21 = dQ_k / dtheta_j

        Args:
            buses: The state of the system buses.
            admittances: The system admittances.
            p_values: The real power mismatch equations.
            unused_q_values: The reactive power mismatch equations.

        Returns:
            The Jacobian matrix J_21.
        """
        j_21 = numpy.zeros((len(buses), len(buses))) * 1j
        for src in buses:
            k = src.number - 1
            v_k = numpy.abs(src.voltage)
            theta_k = numpy.angle(src.voltage)
            p_k = p_values[k]

            for dst in buses:
                j = dst.number - 1
                v_j = numpy.abs(dst.voltage)
                theta_j = numpy.angle(dst.voltage)
                b_kj = admittances[k][j].imag

                if j != k:
                    j_21[k][j] = -v_k * v_j * b_kj * numpy.sin(theta_k - theta_j)
                else:
                    j_21[k][j] = p_k

        return j_21

    def create_j22(self, buses, admittances, unused_p_values, q_values):
        """Creates the upper-left Jacobian matrix, J_22.

            J_22 = dQ_k / dV_j

        Args:
            buses: The state of the system buses.
            admittances: The system admittances.
            unused_p_values: The real power mismatch equations.
            q_values: The reactive power mismatch equations.

        Returns:
            The Jacobian matrix J_22.
        """
        j_22 = numpy.zeros((len(buses), len(buses))) * 1j
        for src in buses:
            k = src.number - 1
            v_k = numpy.abs(src.voltage)
            theta_k = numpy.angle(src.voltage)
            q_k = q_values[k]

            for dst in buses:
                j = dst.number - 1
                v_j = numpy.abs(dst.voltage)
                theta_j = numpy.angle(dst.voltage)
                b_kj = admittances[k][j].imag

                if j != k:
                    j_22[k][j] = -v_k * b_kj * numpy.sin(theta_k - theta_j)
                else:
                    j_22[k][j] = q_k / v_k - b_kj * v_k

        return j_22

    def is_convergent(self, dp_values, dq_values):
        """Checks if the power flow has converged to a solution.

        Args:
            dp_values: The real power mismatches.
            dq_values: The reactive power mismatches,

        Returns:
            True if the power flow has converged to a solution; false otherwise.
        """
        for dp in dp_values:
            if dp >= self.max_mismatch_mw:
                return False

        for dq in dq_values:
            if dq >= self.max_mismatch_mvar:
                return False

        return True
