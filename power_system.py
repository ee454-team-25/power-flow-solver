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


class Bus(collections.namedtuple('Bus', ['number', 'loads', 'generators'])):
    def type(self):
        if self.generators:
            return BusType.PV
        elif self.loads:
            return BusType.PQ

        return BusType.UNKNOWN


class PowerSystem(collections.namedtuple('PowerSystem', ['buses', 'lines'])):
    def admittance_matrix(self):
        matrix = numpy.zeros((len(self.buses), len(self.buses))) * 1j
        for line in self.lines:
            src = line.source - 1
            dst = line.destination - 1

            y_distributed = 1 / line.distributed_impedance
            y_shunt = line.shunt_admittance / 2

            matrix[src][dst] -= y_distributed
            matrix[dst][src] -= y_distributed
            matrix[src][src] += y_distributed + y_shunt
            matrix[dst][dst] += y_distributed + y_shunt

        return matrix
