"""A module containing a power system API.

The main objects in this module are the Line, Bus, and PowerSystem objects. There are one-to-many relationships between
a power system and its buses and lines. Each line connects two buses and there may be many lines between any two buses.

This model is simplified to allow only one load or generator to be attached at each bus. Loads are specified by the
power they consume, while generators are specified by the active power they inject and their controlled voltage.
"""

import collections
import namedlist
import numpy

# A power line connecting two buses.
#
# Args:
#     source: The source bus number.
#     destination: The destination bus number.
#     distributed_impedance: The per-unit distributed impedance of the power line.
#     shunt_admittance: The per-unit shunt admittance of the power line.
#     max_power: The maximum power rating of the line in MVA.
Line = collections.namedtuple('Line',
                              ['source', 'destination', 'distributed_impedance', 'shunt_admittance', 'max_power'])

# A bus in the power system.
#
# This object is a namedlist instead of a namedtuple because the voltage field is mutable and changes at each iteration
# of the load flow analysis.
#
# Args:
#     number: The bus number.
#     active_power_consumed: The per-unit active power consumed at this bus.
#     reactive_power_consumed: The per-unit reactive power consumed at this bus.
#     active_power_injected: The per-unit active power injected at this bus.
#     voltage: The per-unit voltage at this bus.
Bus = namedlist.namedlist('Bus',
                          ['number', 'active_power_consumed', 'reactive_power_consumed', 'active_power_injected',
                           'voltage'])


class PowerSystem(collections.namedtuple('PowerSystem', ['buses', 'lines'])):
    """An object representing a power system.

    Args:
        buses: A list of buses in the system.
        lines: A list of lines in the system.
    """

    def admittance_matrix(self):
        """Computes the admittance matrix for the system.

        Returns:
            The admittance matrix for the system.
        """
        matrix = numpy.zeros((len(self.buses), len(self.buses))) * 1j
        for line in self.lines:
            src = line.source - 1
            dst = line.destination - 1

            y_distributed = 1 / line.distributed_impedance
            y_shunt = line.shunt_admittance

            matrix[src][dst] -= y_distributed
            matrix[dst][src] -= y_distributed
            matrix[src][src] += y_distributed + y_shunt
            matrix[dst][dst] += y_distributed + y_shunt

        return matrix
