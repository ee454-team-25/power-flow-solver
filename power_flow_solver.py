"""A module containing a power flow analysis API.

The main object in this module is the PowerFlowSolver, which takes a power system as input and runs iterations of the
Newton-Raphson method to determine the voltages at each bus, relative to a swing bus.

    system = build_system()
    solver = PowerFlowSolver(system)
    while not solver.has_converged():
        solver.step()
"""

import collections
import enum
import numpy

DEFAULT_SWING_BUS_NUMBER = 1
DEFAULT_MAX_ACTIVE_POWER_ERROR = 0.001
DEFAULT_MAX_REACTIVE_POWER_ERROR = 0.001

# A bus estimate object. This object contains estimated power values.
#
# Args:
#     bus: A reference to the actual bus being estimated.
#     bus_type: The bus type.
#     active_power: The estimated active power injection.
#     reactive_power: The estimated reactive power injection.
#     active_power_error: The difference between the estimated and actual active power injection.
#     reactive_power_error: The difference between the estimated and actual reactive power injection.
_BusEstimate = collections.namedtuple(
    '_BusEstimate', ['bus', 'bus_type', 'active_power', 'reactive_power', 'active_power_error', 'reactive_power_error'])


class _BusType(enum.Enum):
    """Bus type enumerations."""
    UNKNOWN = 0
    SWING = 1
    PV = 2
    PQ = 3


class PowerFlowSolver:
    """A power flow solver object."""

    def __init__(self, system, swing_bus_number=DEFAULT_SWING_BUS_NUMBER,
                 max_active_power_error=DEFAULT_MAX_ACTIVE_POWER_ERROR,
                 max_reactive_power_error=DEFAULT_MAX_REACTIVE_POWER_ERROR):
        """Initializes the power flow solver.

        Args:
            system: The power flow system being analyzed.
            swing_bus_number: The bus designated as the swing bus.
            max_active_power_error: The maximum allowed active power mismatch.
            max_reactive_power_error: The maximum allowed reactive power mismatch.
        """
        self._system = system
        self._swing_bus_number = swing_bus_number
        self._max_active_power_error = max_active_power_error
        self._max_reactive_power_error = max_reactive_power_error

        self._admittance_matrix = self._system.admittance_matrix()
        self._estimates = None
        self._pq_estimates = None

    def has_converged(self):
        """Checks if the analysis has converged to a solution.

        Returns:
            True if the power inject estimates at each bus are equal to the actual power injection (within some
            allowable margin), false otherwise.
        """
        if not self._estimates or not self._pq_estimates:
            return False

        # Check active power mismatches.
        for estimate in self._estimates.values():
            if numpy.abs(estimate.active_power_error) > self._max_active_power_error:
                return False

        # Check reactive power mismatches.
        for estimate in self._pq_estimates.values():
            if numpy.abs(estimate.reactive_power_error) > self._max_reactive_power_error:
                return False

        return True

    def step(self):
        """Executes a step of the power flow analysis using the Newton-Raphson method."""
        self._compute_estimates()
        jacobian = self._jacobian()
        corrections = self._corrections(jacobian)
        self._apply_corrections(corrections)

    def _compute_estimates(self):
        """Computes power injection estimates for each bus and splits out PQ buses."""
        self._estimates = self._bus_power_estimates()
        self._pq_estimates = {i.bus.number: i for i in self._estimates.values() if i.bus_type == _BusType.PQ}

    def _bus_type(self, bus):
        """Classifies a given bus based on which parameters specify it.

        Args:
            bus: The bus to classify.

        Returns:
            The bus classification.
        """
        if bus.number == self._swing_bus_number:
            return _BusType.SWING
        elif bus.active_power_injected:
            return _BusType.PV
        elif bus.active_power_consumed or bus.reactive_power_consumed:
            return _BusType.PQ

        return _BusType.UNKNOWN

    def _bus_power_estimates(self):
        """Computes power injection estimates for each bus.

        Returns:
            A dict mapping a bus number to its power injection estimate.
        """
        estimates = {}
        for src in self._system.buses:
            bus_type = self._bus_type(src)
            if bus_type not in (_BusType.PV, _BusType.PQ):
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
            if bus_type == _BusType.PV:
                p_error += src.active_power_injected
            elif bus_type == _BusType.PQ:
                p_error -= src.active_power_consumed
                q_error -= src.reactive_power_consumed

            estimates[src.number] = _BusEstimate(src, bus_type, p, q, p_error, q_error)

        return estimates

    def _jacobian(self):
        """Computes the Jacobian for the power flow."""
        j11 = self._jacobian_11()
        j12 = self._jacobian_12()
        j21 = self._jacobian_21()
        j22 = self._jacobian_22()
        j1 = numpy.concatenate([j11, j12], axis=1)
        j2 = numpy.concatenate([j21, j22], axis=1)
        return numpy.concatenate([j1, j2], axis=0)

    def _jacobian_11(self):
        """Computes the Jacobian submatrix J11."""
        j11 = numpy.zeros((len(self._estimates), len(self._estimates)))
        for row, src_number in enumerate(self._estimates):
            src = self._estimates[src_number]
            k = src_number - 1
            v_k = numpy.abs(src.bus.voltage)
            theta_k = numpy.angle(src.bus.voltage)
            q_k = src.reactive_power

            for col, dst_number in enumerate(self._estimates):
                dst = self._estimates[dst_number]
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
                    j11[row][col] = -q_k - (v_k ** 2) * b_kj

        return j11

    def _jacobian_12(self):
        """Computes the Jacobian submatrix J12."""
        j12 = numpy.zeros((len(self._estimates), len(self._pq_estimates)))
        for row, src_number in enumerate(self._estimates):
            src = self._estimates[src_number]
            k = src_number - 1
            v_k = numpy.abs(src.bus.voltage)
            theta_k = numpy.angle(src.bus.voltage)
            p_k = src.active_power

            for col, dst_number in enumerate(self._pq_estimates):
                dst = self._pq_estimates[dst_number]
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

    def _jacobian_21(self):
        """Computes the Jacobian submatrix J21."""
        j21 = numpy.zeros((len(self._pq_estimates), len(self._estimates)))
        for row, src_number in enumerate(self._pq_estimates):
            src = self._pq_estimates[src_number]
            k = src_number - 1
            v_k = numpy.abs(src.bus.voltage)
            theta_k = numpy.angle(src.bus.voltage)
            p_k = src.active_power

            for col, dst_number in enumerate(self._estimates):
                dst = self._estimates[dst_number]
                j = dst.bus.number - 1
                v_j = numpy.abs(dst.bus.voltage)
                theta_j = numpy.angle(dst.bus.voltage)
                theta_kj = theta_k - theta_j

                y_kj = self._admittance_matrix[k][j]
                g_kj = y_kj.real
                b_kj = y_kj.imag

                if k != j:
                    j21[row][col] = -v_k * v_j * (g_kj * numpy.cos(theta_kj) + b_kj * numpy.sin(theta_kj))
                else:
                    j21[row][col] = p_k - g_kj * v_k ** 2

        return j21

    def _jacobian_22(self):
        """Computes the Jacobian submatrix J22."""
        j22 = numpy.zeros((len(self._pq_estimates), len(self._pq_estimates)))
        for row, src_number in enumerate(self._pq_estimates):
            src = self._pq_estimates[src_number]
            k = src_number - 1
            v_k = numpy.abs(src.bus.voltage)
            theta_k = numpy.angle(src.bus.voltage)
            q_k = src.reactive_power

            for col, dst_number in enumerate(self._pq_estimates):
                dst = self._pq_estimates[dst_number]
                j = dst.bus.number - 1
                theta_j = numpy.angle(dst.bus.voltage)
                theta_kj = theta_k - theta_j

                y_kj = self._admittance_matrix[k][j]
                g_kj = y_kj.real
                b_kj = y_kj.imag

                if k != j:
                    j22[row][col] = v_k * (g_kj * numpy.sin(theta_kj) - b_kj * numpy.cos(theta_kj))
                else:
                    j22[row][col] = q_k / v_k - b_kj * v_k

        return j22

    def _corrections(self, jacobian):
        """Compute corrective factors to apply to voltage phase angles and magnitudes.

        This method executes an iteration of the Newton-Raphson method. The state vector is given from the list of
        active and reactive power injection mismatches, and the corrective factors are computed by multiplying the
        state vector by the inverse Jacobian.

            dx = J^(-1)x

        There are expected to be phase angle corrections for all PV and PQ buses, and magnitude corrections for all PQ
        buses.

        Args:
            jacobian: The Jacobian matrix for the system.

        Returns:
            An ordered list of voltage phase angle and magnitude corrections.
        """
        p_errors = [i.active_power_error for i in self._estimates.values()]
        q_errors = [i.reactive_power_error for i in self._pq_estimates.values()]
        errors = numpy.transpose([p_errors + q_errors])
        corrections = numpy.matmul(numpy.linalg.inv(jacobian), errors)
        return corrections.transpose()[0]

    def _apply_corrections(self, corrections):
        """Applies a list of voltage corrections to bus estimates.

        Args:
            corrections: A list of voltage phase angle and magnitude corrections.
        """
        angle_corrections = corrections[0:len(self._estimates)]
        magnitude_corrections = corrections[len(self._estimates):]

        for c, e in zip(angle_corrections, self._estimates.values()):
            magnitude = numpy.abs(e.bus.voltage)
            angle = numpy.angle(e.bus.voltage) + c
            e.bus.voltage = magnitude * numpy.exp(1j * angle)

        for c, e in zip(magnitude_corrections, self._pq_estimates.values()):
            magnitude = numpy.abs(e.bus.voltage) + c
            angle = numpy.angle(e.bus.voltage)
            e.bus.voltage = magnitude * numpy.exp(1j * angle)
