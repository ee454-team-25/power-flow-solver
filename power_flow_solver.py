import collections
import enum
import numpy

DEFAULT_SWING_BUS_NUMBER = 1

_BusEstimate = collections.namedtuple('_BusEstimate',
                                      ['bus', 'type', 'active_power', 'reactive_power', 'active_power_error',
                                       'reactive_power_error'])


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
        # jacobian = self._jacobian(self._bus_power_estimates())
        # corrections = self._corrections(jacobian, p_errors, q_errors)
        # self._apply_corrections(corrections)
        pass

    def _bus_type(self, bus):
        if bus.number == self._swing_bus_number:
            return _BusType.SWING
        elif bus.active_power_injected:
            return _BusType.PV
        elif bus.active_power_consumed or bus.reactive_power_consumed:
            return _BusType.PQ

        return _BusType.UNKNOWN

    def _bus_power_estimates(self):
        estimates = {}
        for src in self._system.buses:
            type = self._bus_type(src)
            if type not in (_BusType.PV, _BusType.PQ):
                continue

            p = 0
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

                p -= v_k * v_i * (g_ki * numpy.cos(theta_ki) + b_ki * numpy.sin(theta_ki))
                q -= v_k * v_i * (g_ki * numpy.sin(theta_ki) - b_ki * numpy.cos(theta_ki))

            p_error = p
            q_error = q
            if type == _BusType.PV:
                p_error += src.active_power_injected
            elif type == _BusType.PQ:
                p_error -= src.active_power_consumed
                q_error -= src.reactive_power_consumed

            estimates[src.number] = _BusEstimate(src, type, p, q, p_error, q_error)

        return estimates

    def _jacobian_11(self, estimates):
        j11 = numpy.zeros((len(estimates), len(estimates)))
        for row, src_number in enumerate(estimates):
            src = estimates[src_number]
            k = src_number - 1
            v_k = numpy.abs(src.bus.voltage)
            theta_k = numpy.angle(src.bus.voltage)
            q_k = src.reactive_power

            for col, dst_number in enumerate(estimates):
                dst = estimates[dst_number]
                j = dst.bus.number - 1
                v_j = numpy.abs(dst.bus.voltage)
                theta_j = numpy.angle(dst.bus.voltage)
                theta_kj = theta_k - theta_j

                y_kj = self._admittance_matrix[k][j]
                g_kj = y_kj.real
                b_kj = y_kj.imag

                if k != j:
                    j11[row][col] = v_k * v_j * (g_kj * numpy.sin(theta_kj) - b_kj * numpy.cos(theta_kj))
                else:
                    j11[row][col] = -q_k - v_k ** 2 * b_kj

        return j11

    def _jacobian_12(self, estimates):
        j12 = numpy.zeros((len(estimates), len(estimates)))
        for row, src_number in enumerate(estimates):
            src = estimates[src_number]
            k = src_number - 1
            v_k = numpy.abs(src.bus.voltage)
            theta_k = numpy.angle(src.bus.voltage)
            p_k = src.active_power

            for col, dst_number in enumerate(estimates):
                dst = estimates[dst_number]
                j = dst.bus.number - 1
                theta_j = numpy.angle(dst.bus.voltage)
                theta_kj = theta_k - theta_j

                y_kj = self._admittance_matrix[k][j]
                g_kj = y_kj.real
                b_kj = y_kj.imag

                if k != j:
                    j12[row][col] = v_k * (g_kj * numpy.cos(theta_kj) + b_kj * numpy.sin(theta_kj))
                else:
                    j12[row][col] = p_k / v_k + g_kj * v_k

        return j12
