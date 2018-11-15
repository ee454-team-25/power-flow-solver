import collections
import numpy

Load = collections.namedtuple('Load', ['active_power', 'reactive_power'])

Generator = collections.namedtuple('Generator', ['voltage', 'active_power'])

Bus = collections.namedtuple('Bus', ['number', 'loads', 'generators'])

Line = collections.namedtuple('Line',
                              ['source', 'destination', 'distributed_impedance', 'shunt_admittance', 'max_power'])


class PowerSystem:
    def __init__(self, buses, lines):
        self._buses = buses
        self._lines = lines

    @property
    def buses(self):
        return self._buses

    @property
    def lines(self):
        return self._lines

    def admittance_matrix(self):
        matrix = numpy.zeros((len(self._buses), len(self._buses))) * 1j
        for line in self._lines:
            src = line.source - 1
            dst = line.destination - 1

            y_distributed = 1 / line.distributed_impedance
            y_shunt = line.shunt_admittance / 2

            matrix[src][dst] -= y_distributed
            matrix[dst][src] -= y_distributed
            matrix[src][src] += y_distributed + y_shunt
            matrix[dst][dst] += y_distributed + y_shunt

        return matrix
