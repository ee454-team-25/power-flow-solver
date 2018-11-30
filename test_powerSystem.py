import numpy
import power_system
import power_system_builder
import unittest


class TestPowerSystem(unittest.TestCase):
    def test_admittance_matrix(self):
        filename = 'data/Data.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()

        actual = system.admittance_matrix()
        expected = [
            [6.03 - 19.45j, -5 + 15.26j, 0, 0, -1.03 + 4.23j, 0, 0, 0, 0, 0, 0, 0],
            [-5 + 15.26j, 9.52 - 30.27j, -1.14 + 4.78j, -1.69 + 5.12j, -1.7 + 5.19j, 0, 0, 0, 0, 0, 0, 0],
            [0, -1.14 + 4.78j, 2.54 - 5.11j, -1.4 + 0.36j, 0, 0, 0, 0, 0, 0, 0, 0],
            [0, -1.69 + 5.12j, -1.4 + 0.36j, 9.93 - 28.83j, -6.84 + 21.58j, 0, 1.8j, 0, 0, 0, 0, 0],
            [-1.03 + 4.23j, -1.7 + 5.19j, 0, -6.84 + 21.58j, 9.57 - 34.93j, 3.97j, 0, 0, 0, 0, 0, 0],
            [0, 0, 0, 0, 3.97j, 6.58 - 17.34j, 0, 0, -1.96 + 4.09j, -1.53 + 3.18j, -3.10 + 6.10j, 0],
            [0, 0, 0, 1.8j, 0, 0, 5.33 - 15.19j, -3.9 + 10.37j, 0, 0, 0, -1.42 + 3.03j],
            [0, 0, 0, 0, 0, 0, -3.9 + 10.37j, 5.78 - 14.77j, -1.88 + 4.4j, 0, 0, 0],
            [0, 0, 0, 0, 0, -1.96 + 4.09j, 0, -1.88 + 4.4j, 3.84 - 8.5j, 0, 0, 0],
            [0, 0, 0, 0, 0, -1.53 + 3.18j, 0, 0, 0, 4.01 - 5.43j, -2.49 + 2.25j, 0],
            [0, 0, 0, 0, 0, -3.1 + 6.1j, 0, 0, 0, -2.49 + 2.25j, 6.72 - 10.67j, -1.14 + 2.31j],
            [0, 0, 0, 0, 0, 0, -1.42 + 3.03j, 0, 0, 0, -1.14 + 2.31j, 2.56 - 5.34j]]

        numpy.testing.assert_almost_equal(actual, expected, 2)

    def test_admittance_matrix_powell(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()

        actual = system.admittance_matrix()
        expected = numpy.array([
            [10.958904 - 25.997397j, -3.424658 + 7.534247j, -3.424658 + 7.534247j, 0j, -4.109589 + 10.958904j],
            [-3.424658 + 7.534247j, 11.672080 - 26.060948j, -4.123711 + 9.278351j, 0j, -4.123711 + 9.278351j],
            [-3.424658 + 7.534247j, -4.123711 + 9.278351j, 10.475198 - 23.119061j, -2.926829 + 6.341463j, 0j],
            [0j, 0j, -2.926829 + 6.341463j, 7.050541 - 15.594814j, -4.123711 + 9.278351j],
            [-4.109589 + 10.958904j, -4.123711 + 9.278351j, 0j, -4.123711 + 9.278351j, 12.357012 - 29.485605j]])

        numpy.testing.assert_almost_equal(expected, actual, 6)

    def test_admittance_matrix_nptel(self):
        filename = 'data/Sample-nptel.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()

        actual = system.admittance_matrix()
        expected = numpy.array([
            [2.6923 - 13.4115j, -1.9231 + 9.6154j, 0, 0, -0.7692 + 3.8462j],
            [-1.9231 + 9.6154j, 3.6538 - 18.1942j, -0.9615 + 4.8077j, 0, -0.7692 + 3.8462j],
            [0, -0.9615 + 4.8077j, 2.2115 - 11.0027j, -0.7692 + 3.8462j, -0.4808 + 2.4038j],
            [0, 0, -0.7692 + 3.8462j, 1.1538 - 5.6742j, -0.3846 + 1.9231j],
            [-0.7692 + 3.8462j, -0.7692 + 3.8462j, -0.4808 + 2.4038j, -0.3846 + 1.9231j, 2.4038 - 11.8942j]])

        numpy.testing.assert_almost_equal(actual, expected, 4)
