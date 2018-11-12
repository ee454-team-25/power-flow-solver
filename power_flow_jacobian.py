"""A module that creates the Jacobian matrix for a power system.

For each load bus there are two estimates (for P and Q), and for each generator bus there is one estimate (for P). The
slack bus is considered a reference bus, so no equations are derived from it. Let m be the number of load buses, and n
be the number of generator buses. Then there are (m + n) active power equations, m reactive power equations, and
2m + n total equations.

For each load bus there are two unknowns (voltage magnitude and angle), and for each generator bus there is one unknown
voltage angle). Then there are m unknown voltage angles and (n + m) unknown voltage magnitudes.

The Jacobian is a real-valued, 4 by 4 matrix with the following form:

    J = [[J_11] [J_12]]
        [[J_21] [J_22]]

Let there be m load buses, n generator buses, and let B[k][j] be the susceptance from bus k to j. Then the Jacobian
submatrices have the following dimensions and definitions:

    J_11: (m + n) by (m + n)
    j_11[k][j] = dP[k] / dtheta[j] = -B[k][j] V[k] V[j] cos(theta[k] - theta[j]) if j != k
                                   = -Q[k] - V[k]^2 B[k][j]                      if j == k

    J_12: (m + n) by m
    j_12[k][j] = dP[k] / dV[j]     = B[k][j] V[k] sin(theta[k] - theta[j])       if j != k
                                   = P[k] / V[k]                                 if j == k

    J_21: m by (m + n)
    j_21[k][j] = dQ[k] / dtheta[j] = -B[k][j] V[k] V[j] sin(theta[k] - theta[j]) if j != k
                                   = P[k]                                        if j == k

    J_22: m by m
    j_22[k][j] = dQ[k] / dV[j]     = -B[k][j] V[k] cos(theta[k] - theta[j])      if j != k
                                   = Q[k] / V[k] - B[k][j] V[j]                  if j == k
"""

import numpy
import operator
import power_flow


def create_jacobian_matrix(bus_estimates, admittances):
    """Creates the Jacobian matrix for the state of a power system.

    Args:
        bus_estimates: The state of the system buses.
        admittances: The system admittances.

    Returns:
        The Jacobian matrix for the power system.
    """
    # Split load and generator buses.
    load_bus_estimates = [i for i in bus_estimates.values() if i.type == power_flow.BusType.LOAD]
    load_bus_estimates = sorted(load_bus_estimates, key=operator.attrgetter('number'))
    generator_bus_estimates = [i for i in bus_estimates.values() if i.type == power_flow.BusType.GENERATOR]
    generator_bus_estimates = sorted(generator_bus_estimates, key=operator.attrgetter('number'))

    # Create submatrices.
    j_11 = create_j11_matrix(load_bus_estimates, generator_bus_estimates, admittances)
    j_12 = create_j12_matrix(load_bus_estimates, generator_bus_estimates, admittances)
    j_21 = create_j21_matrix(load_bus_estimates, generator_bus_estimates, admittances)
    j_22 = create_j22_matrix(load_bus_estimates, admittances)

    # Combine matrices together.
    j_1 = numpy.concatenate([j_11, j_12], axis=1)
    j_2 = numpy.concatenate([j_21, j_22], axis=1)
    j = numpy.concatenate([j_1, j_2], axis=0)
    return j


def create_j11_matrix(load_bus_estimates, generator_bus_estimates, admittances):
    """Creates the upper-left Jacobian matrix, J_11.

    Args:
        bus_estimates: The state of the system buses.
        admittances: The system admittances.

    Returns:
        The (m + n) by (m + n) matrix J_11.
    """
    bus_estimates = load_bus_estimates + generator_bus_estimates
    j_11 = numpy.zeros((len(bus_estimates), len(bus_estimates)))
    for x, src in enumerate(bus_estimates):
        k = src.number - 1
        v_k = numpy.abs(src.voltage)
        theta_k = numpy.angle(src.voltage)
        q_k = src.power.imag

        for y, dst in enumerate(bus_estimates):
            j = dst.number - 1
            v_j = numpy.abs(dst.voltage)
            theta_j = numpy.angle(dst.voltage)
            b_kj = admittances[k][j].imag

            if j != k:
                j_11[x][y] = -v_k * v_j * b_kj * numpy.cos(theta_k - theta_j)
            else:
                j_11[x][y] = -q_k - v_k ** 2 * b_kj

    return j_11


def create_j12_matrix(load_bus_estimates, generator_bus_estimates, admittances):
    """Creates the upper-left Jacobian matrix, J_12.

    Args:
        bus_estimates: The state of the system buses.
        admittances: The system admittances.

    Returns:
        The (m + n) by m matrix J_12.
    """
    bus_estimates = load_bus_estimates + generator_bus_estimates
    j_12 = numpy.zeros((len(bus_estimates), len(load_bus_estimates)))
    for x, src in enumerate(bus_estimates):
        k = src.number - 1
        v_k = numpy.abs(src.voltage)
        theta_k = numpy.angle(src.voltage)
        p_k = src.power.real

        for y, dst in enumerate(load_bus_estimates):
            j = dst.number - 1
            theta_j = numpy.angle(dst.voltage)
            b_kj = admittances[k][j].imag

            if j != k:
                j_12[x][y] = v_k * b_kj * numpy.sin(theta_k - theta_j)
            else:
                j_12[x][y] = p_k / v_k

    return j_12


def create_j21_matrix(load_bus_estimates, generator_bus_estimates, admittances):
    """Creates the upper-left Jacobian matrix, J_21.

    Args:
        load_bus_estimates: The states of the system load buses.
        generator_bus_estimates: The states of the system generator buses.
        admittances: The system admittances.

    Returns:
        The m by (m + n) matrix J_21.
    """
    bus_estimates = load_bus_estimates + generator_bus_estimates
    j_21 = numpy.zeros((len(load_bus_estimates), len(bus_estimates)))
    for x, src in enumerate(load_bus_estimates):
        k = src.number - 1
        v_k = numpy.abs(src.voltage)
        theta_k = numpy.angle(src.voltage)
        p_k = src.power.real

        for y, dst in enumerate(bus_estimates):
            j = dst.number - 1
            v_j = numpy.abs(dst.voltage)
            theta_j = numpy.angle(dst.voltage)
            b_kj = admittances[k][j].imag

            if j != k:
                j_21[x][y] = -v_k * v_j * b_kj * numpy.sin(theta_k - theta_j)
            else:
                j_21[x][y] = p_k

    return j_21


def create_j22_matrix(load_bus_estimates, admittances):
    """Creates the upper-left Jacobian matrix, J_22.

    Args:
        load_bus_estimates: The state of the system buses.
        admittances: The system admittances.

    Returns:
        The Jacobian matrix J_22.
    """
    j_22 = numpy.zeros((len(load_bus_estimates), len(load_bus_estimates)))
    for x, src in enumerate(load_bus_estimates):
        k = src.number - 1
        v_k = numpy.abs(src.voltage)
        theta_k = numpy.angle(src.voltage)
        q_k = src.power.imag

        for y, dst in enumerate(load_bus_estimates):
            j = dst.number - 1
            theta_j = numpy.angle(dst.voltage)
            b_kj = admittances[k][j].imag

            if j != k:
                j_22[x][y] = -v_k * b_kj * numpy.cos(theta_k - theta_j)
            else:
                j_22[x][y] = q_k / v_k - b_kj * v_k

    return j_22
