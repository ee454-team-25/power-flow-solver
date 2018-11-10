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
        self, bus_data, line_data, slack_bus, start_voltage, power_base_mva, max_mismatch_mw, max_mismatch_mvar):
        """Creates a new power flow object.

        Args:
            bus_data: The openpyxl worksheet containing bus data.
            line_data: The openpyxl worksheet containing line data.
            slack_bus: The slack bus number.
            start_voltage: The initial voltages to use at each bus.
            power_base_mva: The power base in MVA.
            max_mismatch_mw: The maximum allowable real power mismatch for a solution to be considered final.
            max_mismatch_mvar: The maximum allowable reactive power mismatch for a solution to be considered final.
        """
        self.bus_data = bus_data
        self.line_data = line_data
        self.slack_bus = slack_bus
        self.start_voltage = start_voltage
        self.power_base_mva = power_base_mva
        self.max_mismatch_mw = max_mismatch_mw
        self.max_mismatch_mvar = max_mismatch_mvar

    def execute(self):
        """Executes a power flow analysis.

        Returns:
            An array of buses and their steady-state solutions.
        """
        # Read bus data.
        buses = self.read_bus_data()

        # Read admittance matrix.
        # admittances = self.read_admittance_matrix(len(buses))

        # TODO(kjiwa): Apply the Newton-Raphson method until a solution is found.

        # Create power flow equations.
        # mismatches = self.create_mismatch_equations(buses, admittances)

        # Check for convergence.
        # if self.is_convergent(mismatches):
        #     break

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
            An array containing load bus data.
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
        buses[self.slack_bus - 1].s = -s_total
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

    def create_mismatch_equations(self, buses, admittances):
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
        dS = []
        for src in buses:
            P = -src.p
            Q = -src.q

            k = src.number - 1
            v_k = numpy.abs(src.voltage)
            theta_k = numpy.angle(src.voltage)

            for dst in buses:
                i = dst.number - 1
                v_i = numpy.abs(dst.voltage)
                theta_i = numpy.angle(dst.voltage)
                b_ki = admittances[k][i]

                P += v_k * v_i * b_ki * numpy.sin(theta_k - theta_i)
                Q += -v_k * v_i * b_ki * numpy.cos(theta_k - theta_i)

            dS.append(P + 1j * Q)

        return dS

    def is_convergent(self, mismatches):
        """Checks if the power flow has converged to a solution.

        Args:
            mismatches: Bus power mismatches.

        Returns:
            True if the power flow has converged to a solution; false otherwise.
        """
        convergent = True
        for mismatch in mismatches:
            if mismatch.real >= self.max_mismatch_mw or mismatch.imag >= self.max_mismatch_mvar:
                convergent = False
                break

        return convergent
