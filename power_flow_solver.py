import copy
import enum
import numpy

DEFAULT_SWING_BUS_NUMBER = 1


class _BusType(enum.Enum):
    UNKNOWN = 0
    SWING = 1
    PV = 2
    PQ = 3


class PowerFlowSolver:
    def __init__(self, system, swing_bus_number=DEFAULT_SWING_BUS_NUMBER):
        self._system = system
        self._swing_bus_number = swing_bus_number
        self._has_converged = False
        self._admittance_matrix = self._system.admittance_matrix()

    @property
    def has_converged(self):
        return self._has_converged

    def step(self):
        # p_estimates = self._p_estimates()
        # q_estimates = self._q_estimates()
        # p_errors = self._p_errors(p_estimates)
        # q_errors = self._q_errors(q_estimates)
        pass

    def _bus_type(self, bus):
        if bus.number == self._swing_bus_number:
            return _BusType.SWING
        elif bus.active_power_injected:
            return _BusType.PV
        elif bus.active_power_consumed:
            return _BusType.PQ

        return _BusType.UNKNOWN

    def _p_estimates(self):
        result = []
        for src in self._system.buses:
            type = self._bus_type(src)
            if type not in (_BusType.PV, _BusType.PQ):
                continue

            p = 0
            v_k = numpy.abs(src.voltage)
            theta_k = numpy.angle(src.voltage)

            for dst in self._system.buses:
                v_i = numpy.abs(dst.voltage)
                theta_i = numpy.angle(dst.voltage)

                y_ki = self._admittance_matrix[src.number - 1][dst.number - 1]
                g_ki = y_ki.real
                b_ki = y_ki.imag
                theta_ki = theta_k - theta_i

                p -= v_k * v_i * (g_ki * numpy.cos(theta_ki) + b_ki * numpy.sin(theta_ki))

            result.append(p)

        return result

    def _q_estimates(self):
        result = []
        for src in self._system.buses:
            if self._bus_type(src) != _BusType.PQ:
                continue

            q = 0
            v_k = numpy.abs(src.voltage)
            theta_k = numpy.angle(src.voltage)

            for dst in self._system.buses:
                v_i = numpy.abs(dst.voltage)
                theta_i = numpy.angle(dst.voltage)

                y_ki = self._admittance_matrix[src.number - 1][dst.number - 1]
                g_ki = y_ki.real
                b_ki = y_ki.imag
                theta_ki = theta_k - theta_i

                q -= v_k * v_i * (g_ki * numpy.sin(theta_ki) - b_ki * numpy.cos(theta_ki))

            result.append(q)

        return result

    def _p_errors(self, p_estimates):
        errors = copy.deepcopy(p_estimates)
        index = 0
        for bus in self._system.buses:
            type = self._bus_type(bus)
            if type == _BusType.PV:
                errors[index] += bus.active_power_injected
                index += 1
            elif type == _BusType.PQ:
                errors[index] -= bus.active_power_consumed
                index += 1

        return errors

    def _q_errors(self, q_estimates):
        errors = copy.deepcopy(q_estimates)
        index = 0
        for bus in self._system.buses:
            if self._bus_type(bus) != _BusType.PQ:
                continue

            errors[index] -= bus.reactive_power_consumed
            index += 1

        return errors
