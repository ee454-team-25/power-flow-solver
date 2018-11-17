import numpy
import power_flow_solver
import power_system_builder
import unittest


class TestPowerFlowSolver(unittest.TestCase):
    @staticmethod
    def build_solver():
        builder = power_system_builder.ExcelPowerSystemBuilder('data/Sample-Powell-3.1.xlsx')
        return power_flow_solver.PowerFlowSolver(builder.build_system())

    def test_errors(self):
        solver = TestPowerFlowSolver.build_solver()

        p_expected = [-0.4, -0.25, -0.4, -0.5]
        q_expected = [-0.17, -0.115, -0.175, -0.17]

        actual = solver._bus_power_estimates()
        p_actual = [estimate.active_power_error for estimate in actual.values()
                    if estimate.type in (power_flow_solver._BusType.PQ, power_flow_solver._BusType.PV)]
        q_actual = [estimate.reactive_power_error for estimate in actual.values()
                    if estimate.type == power_flow_solver._BusType.PQ]

        numpy.testing.assert_array_almost_equal(p_expected, p_actual)
        numpy.testing.assert_array_almost_equal(q_expected, q_actual)

    def test_jacobian_11(self):
        solver = TestPowerFlowSolver.build_solver()
        solver.compute_estimates()

        expected = [[26.030948, -9.278351, 0, -9.278350],
                    [-9.278351, 23.084061, -6.341463, 0],
                    [0, -6.341463, 15.569814, -9.278351],
                    [-9.278351, 0, -9.278351, 29.455605]]
        actual = solver._jacobian_11()
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian_12(self):
        solver = TestPowerFlowSolver.build_solver()
        solver.compute_estimates()

        expected = [[11.672080, -4.123711, 0, -4.123711],
                    [-4.123711, 10.475198, -2.926829, 0],
                    [0, -2.926829, 7.050541, -4.123711],
                    [-4.123711, 0, -4.123711, 12.357011]]

        actual = solver._jacobian_12()
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian_21(self):
        solver = TestPowerFlowSolver.build_solver()
        solver.compute_estimates()

        expected = [[-11.672080, 4.123711, 0, 4.123711],
                    [4.123711, -10.475198, 2.926829, 0],
                    [0, 2.926829, -7.050541, 4.123711],
                    [4.123711, 0, 4.123711, -12.357011]]
        actual = solver._jacobian_21()
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian_22(self):
        solver = TestPowerFlowSolver.build_solver()
        solver.compute_estimates()

        expected = [[26.090948, -9.278351, 0, -9.278351],
                    [-9.278351, 23.154061, -6.341463, 0],
                    [0, -6.341463, 15.619814, -9.278351],
                    [-9.278351, 0, -9.278351, 29.515605]]
        actual = solver._jacobian_22()
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian(self):
        solver = TestPowerFlowSolver.build_solver()
        solver.compute_estimates()

        expected = [[26.030948, -9.278351, 0, -9.278350, 11.672080, -4.123711, 0, -4.123711],
                    [-9.278351, 23.084061, -6.341463, 0, -4.123711, 10.475198, -2.926829, 0],
                    [0, -6.341463, 15.569814, -9.278351, 0, -2.926829, 7.050541, -4.123711],
                    [-9.278351, 0, -9.278351, 29.455605, -4.123711, 0, -4.123711, 12.357011],
                    [-11.672080, 4.123711, 0, 4.123711, 26.090948, -9.278351, 0, -9.278351],
                    [4.123711, -10.475198, 2.926829, 0, -9.278351, 23.154061, -6.341463, 0],
                    [0, 2.926829, -7.050541, 4.123711, 0, -6.341463, 15.619814, -9.278351],
                    [4.123711, 0, 4.123711, -12.357011, -9.278351, 0, -9.278351, 29.515605]]
        actual = solver._jacobian()
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_corrections(self):
        solver = TestPowerFlowSolver.build_solver()
        solver.compute_estimates()

        expected = numpy.transpose(
            [[-0.040599, -0.040003, -0.060281, -0.04515, -0.041184, -0.041492, -0.061248, -0.042768]])
        actual = solver._corrections(solver._jacobian())
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_step(self):
        solver = TestPowerFlowSolver.build_solver()
        solver.step()

        expected_magnitudes = [1, 0.958729, 0.958447, 0.938681, 0.957170]
        actual_magnitudes = [numpy.abs(i.voltage) for i in solver._system.buses]
        numpy.testing.assert_array_almost_equal(expected_magnitudes, actual_magnitudes, 3)

        expected_angles = [0, -0.040160, -0.039524, -0.059629, -0.04515]
        actual_angles = [numpy.angle(i.voltage) for i in solver._system.buses]
        numpy.testing.assert_array_almost_equal(expected_angles, actual_angles, 3)
