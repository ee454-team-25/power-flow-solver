"""A module containing power system reporters.

Each reporter analyzes the power system and produces a tabular report as a string.
"""

import numpy
import tabulate

TABULATE_FLOAT_FMT = '.4f'


class BusVoltageReporter:
    """A reporter for bus voltages."""

    def __init__(self, system):
        """Creates a new reporter.

        Args:
            system: The power system being analyzed.
        """
        self._system = system

    def report(self):
        """Reports the voltage at each bus."""
        headers = ['Bus', 'Voltage (V)', 'Phase (deg)']
        table = []
        for bus in self._system.buses:
            table.append([bus.number, numpy.abs(bus.voltage), numpy.rad2deg(numpy.angle(bus.voltage))])

        return tabulate.tabulate(table, headers=headers, floatfmt=TABULATE_FLOAT_FMT)


class LinePowerReporter:
    """A reporter for transmission line power."""

    def __init__(self, system, power_base):
        """Creates a new reporter.

        Args:
            system: The power system being analyzed.
            power_base: The power base in MVA.
        """
        self._system = system
        self._power_base = power_base

    def report(self):
        """Reports the active, reactive, and apparent power for each transmission line.

        TODO(kjiwa): Report whether a line has exceeded its power rating.
        """
        headers = ['Line', 'Active Power (MW)', 'Reactive Power (Mvar)', 'Apparent Power (MVA)']
        table = []
        for line in self._system.lines:
            src = self._system.buses[line.source - 1]
            dst = self._system.buses[line.destination - 1]

            v = dst.voltage - src.voltage
            complex_power = self._power_base * numpy.abs(v) ** 2 / numpy.conj(line.distributed_impedance)
            line_name = '{}-{}'.format(src.number, dst.number)
            table.append([line_name, complex_power.real, complex_power.imag, numpy.abs(complex_power)])

        return tabulate.tabulate(table, headers=headers, floatfmt=TABULATE_FLOAT_FMT)
