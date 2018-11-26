import numpy
import power_system
import power_system_builder
import unittest


class TestPowerSystem(unittest.TestCase):
    def test_admittance_matrix(self):
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

        numpy.testing.assert_almost_equal(expected, actual, 4)
