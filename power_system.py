"""A module containing a power system API.

The main objects in this module are the Line, Bus, and PowerSystem objects. There are one-to-many relationships between
a power system and its buses and lines. Each line connects two buses and there may be many lines between any two buses.

This model is simplified to allow only one load or generator to be attached at each bus. Loads are specified by the
power they consume, while generators are specified by the active power they inject and their controlled voltage.
"""

import dataclasses
import numpy
import typing


@dataclasses.dataclass()
class Bus:
    """A bus in the power system."""
    number: int
    active_power_consumed: float
    reactive_power_consumed: float
    active_power_generated: float
    voltage: complex


@dataclasses.dataclass(frozen=True)
class Line:
    """A power line connecting two buses."""
    source: int
    destination: int
    distributed_impedance: complex
    shunt_admittance: complex
    max_power: typing.Optional[float]


@dataclasses.dataclass(frozen=True)
class PowerSystem:
    """An object representing a power system."""
    buses: typing.List[Bus]
    lines: typing.List[Line]

    def admittance_matrix(self):
        """Computes the admittance matrix for the system.

        Returns:
            The admittance matrix for the system.
        """
        matrix = numpy.zeros((len(self.buses), len(self.buses))) * 1j
        for line in self.lines:
            src = [index for index, bus in enumerate(self.buses) if bus.number == line.source][0]
            dst = [index for index, bus in enumerate(self.buses) if bus.number == line.destination][0]

            y_distributed = 1 / line.distributed_impedance
            y_shunt = line.shunt_admittance

            matrix[src][dst] -= y_distributed
            matrix[dst][src] -= y_distributed
            matrix[src][src] += y_distributed + y_shunt
            matrix[dst][dst] += y_distributed + y_shunt

        return matrix
