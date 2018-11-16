import numpy
import power_flow_solver
import power_system_builder
import unittest


class TestPowerFlowSolver(unittest.TestCase):
    def test_p_errors(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()
        solver = power_flow_solver.PowerFlowSolver(system)

        expected = [-0.4, -0.25, -0.4, -0.5]
        actual = solver._p_errors(solver._p_estimates())
        self.assertListEqual(expected, actual)

    def test_q_errors(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()
        solver = power_flow_solver.PowerFlowSolver(system)

        expected = [-0.17, -0.115, -0.175, -0.17]
        actual = solver._q_errors(solver._q_estimates())
        numpy.testing.assert_array_almost_equal(expected, actual)
