import collections
import enum
import numpy

Load = collections.namedtuple('Load', ['active_power', 'reactive_power'])

Generator = collections.namedtuple('Generator', ['voltage', 'active_power'])

Line = collections.namedtuple('Line',
                              ['source', 'destination', 'distributed_impedance', 'shunt_admittance', 'max_power'])


class BusType(enum.Enum):
    UNKNOWN = 0
    PV = 1
    PQ = 2


class Bus:
    def __init__(self, number, loads, generators):
        self._number = number
        self._loads = loads
        self._generators = generators

    def __eq__(self, other):
        if self.number != other.number:
            return False

        for a, b in zip(self.loads, other.loads):
            if a.active_power != b.active_power or a.reactive_power != b.reactive_power:
                return False

        for a, b in zip(self.generators, other.generators):
            if a.voltage != b.voltage or a.active_power != b.active_power:
                return False

        return True

    @property
    def number(self):
        return self._number

    @property
    def loads(self):
        return self._loads

    @property
    def generators(self):
        return self._generators

    def type(self):
        if self._generators:
            return BusType.PV
        elif self._loads:
            return BusType.PQ

        return BusType.UNKNOWN


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
