import numpy
import power_flow_solver
import power_system_builder
import unittest


class TestPowerFlowSolver(unittest.TestCase):
    @staticmethod
    def build_solver(filename):
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        return power_flow_solver.PowerFlowSolver(builder.build_system())

    def test_errors(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-Powell-3.1.xlsx')

        p_expected = [-0.4, -0.25, -0.4, -0.5]
        q_expected = [-0.17, -0.115, -0.175, -0.17]

        actual = solver._bus_power_estimates()
        p_actual = [estimate.active_power_error for estimate in actual.values()
                    if estimate.bus_type in (power_flow_solver.BusType.PQ, power_flow_solver.BusType.PV)]
        q_actual = [estimate.reactive_power_error for estimate in actual.values()
                    if estimate.bus_type == power_flow_solver.BusType.PQ]

        numpy.testing.assert_array_almost_equal(p_expected, p_actual)
        numpy.testing.assert_array_almost_equal(q_expected, q_actual)

    def test_jacobian_11(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-Powell-3.1.xlsx')
        solver._compute_estimates()

        expected = [[26.030948, -9.278351, 0, -9.278350],
                    [-9.278351, 23.084061, -6.341463, 0],
                    [0, -6.341463, 15.569814, -9.278351],
                    [-9.278351, 0, -9.278351, 29.455605]]
        actual = solver._jacobian_11()
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian_12(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-Powell-3.1.xlsx')
        solver._compute_estimates()

        expected = [[11.672080, -4.123711, 0, -4.123711],
                    [-4.123711, 10.475198, -2.926829, 0],
                    [0, -2.926829, 7.050541, -4.123711],
                    [-4.123711, 0, -4.123711, 12.357011]]

        actual = solver._jacobian_12()
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian_21(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-Powell-3.1.xlsx')
        solver._compute_estimates()

        expected = [[-11.672080, 4.123711, 0, 4.123711],
                    [4.123711, -10.475198, 2.926829, 0],
                    [0, 2.926829, -7.050541, 4.123711],
                    [4.123711, 0, 4.123711, -12.357011]]
        actual = solver._jacobian_21()
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian_22(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-Powell-3.1.xlsx')
        solver._compute_estimates()

        expected = [[26.090948, -9.278351, 0, -9.278351],
                    [-9.278351, 23.154061, -6.341463, 0],
                    [0, -6.341463, 15.619814, -9.278351],
                    [-9.278351, 0, -9.278351, 29.515605]]
        actual = solver._jacobian_22()
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-Powell-3.1.xlsx')
        solver._compute_estimates()

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
        solver = TestPowerFlowSolver.build_solver('data/Sample-Powell-3.1.xlsx')
        solver._compute_estimates()

        expected = [-0.040599, -0.040003, -0.060281, -0.04515, -0.041184, -0.041492, -0.061248, -0.042768]
        actual = solver._corrections(solver._jacobian())
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_step(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-Powell-3.1.xlsx')
        solver.step()

        expected_magnitudes = [1, 0.958729, 0.958447, 0.938681, 0.957170]
        actual_magnitudes = [numpy.abs(i.voltage) for i in solver._system.buses]
        numpy.testing.assert_array_almost_equal(expected_magnitudes, actual_magnitudes, 3)

        expected_angles = [0, -0.040160, -0.039524, -0.059629, -0.04515]
        actual_angles = [numpy.angle(i.voltage) for i in solver._system.buses]
        numpy.testing.assert_array_almost_equal(expected_angles, actual_angles, 3)

    def test_convergence(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-Powell-3.1.xlsx')
        self.assertFalse(solver.has_converged())

        for _ in range(0, 3):
            solver.step()
            self.assertFalse(solver.has_converged())

        solver.step()
        self.assertTrue(solver.has_converged())

    def test_jacobian_11_nptel(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-nptel.xlsx')
        solver._compute_estimates()

        expected = [[18.8269, -4.8077, 0, -3.9231],
                    [-4.8077, 11.1058, -3.8462, -2.4519],
                    [0, -3.8462, 5.8077, -1.9615],
                    [-3.9231, -2.4519, -1.9615, 12.4558]]
        actual = solver._jacobian_11()
        numpy.testing.assert_array_almost_equal(expected, actual, 4)

    def test_jacobian_12_nptel(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-nptel.xlsx')
        solver._compute_estimates()

        expected = [[3.5423, -0.9615, 0],
                    [-0.9615, 2.2019, -0.7692],
                    [0, -0.7692, 1.1462],
                    [0.7846, -0.4904, -0.3923]]

        actual = solver._jacobian_12()
        numpy.testing.assert_array_almost_equal(expected, actual, 4)

    def test_jacobian_21_nptel(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-nptel.xlsx')
        solver._compute_estimates()

        expected = [[-3.7654, 0.9615, 0, 0.7846],
                    [0.9615, -2.2212, 0.7692, 0.4904],
                    [0, 0.7692, -1.1615, 0.3923]]
        actual = solver._jacobian_21()
        numpy.testing.assert_array_almost_equal(expected, actual, 4)

    def test_jacobian_22_nptel(self):
        solver = TestPowerFlowSolver.build_solver('data/Sample-nptel.xlsx')
        solver._compute_estimates()

        expected = [[17.5615, -4.8077, 0],
                    [-4.8077, 10.8996, -3.8462],
                    [0, -3.8462, 5.5408]]
        actual = solver._jacobian_22()
        numpy.testing.assert_array_almost_equal(expected, actual, 4)
