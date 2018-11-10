"""A module that creates the Jacobian matrix for a power system.

For an n-bus system, the Jacobian is a 4 by 4 matrix with the following form:

    J = [[J_11] [J_12]]
        [[J_21] [J_22]]

Each submatrix has dimensions (n - 1) by (n - 1). This is because the slack bus is ignored during Jacobian
computations. This gives the resulting matrix dimensions (2n - 2) by (2n - 2). The submatrices have the following
definitions:

    j_11[k][j] = dP[k] / dtheta[j] = -B[k][j] V[k] V[j] sin(theta[k] - theta[j]) if j != k
                                   = -Q[k] - V[k]^2 B[k][j]                      if j == k

    j_12[k][j] = dP[k] / dV[j]     = B[k][j] V[k] sin(theta[k] - theta[j])       if j != k
                                   = P[k] / V[k]                                 if j == k

    j_21[k][j] = dQ[k] / dtheta[j] = -B[k][j] V[k] V[j] sin(theta[k] - theta[j]) if j != k
                                   = P[k]                                        if j == k

    j_22[k][j] = dQ[k] / dV[j]     = -B[k][j] V[k] cos(theta[k] - theta[j])      if j != k
                                   = Q[k] / V[k]                                 if j == k

where B[k][j] is the susceptance from bus k to j.
"""

import numpy


def create_jacobian_matrix(buses, admittances, s_values, slack_bus_number):
    """Creates the Jacobian matrix for the state of a power system.

    Args:
        buses: The state of the system buses.
        admittances: The system admittances.
        s_values: The power mismatch equations.
        slack_bus_number: The slack bus number.

    Returns:
        The Jacobian matrix for the power system.
    """
    # Create submatrices.
    j_11 = create_j11_matrix(buses, admittances, s_values)
    j_12 = create_j12_matrix(buses, admittances, s_values)
    j_21 = create_j21_matrix(buses, admittances, s_values)
    j_22 = create_j22_matrix(buses, admittances, s_values)

    # Remove slack bus entries.
    j_11 = numpy.delete(j_11, slack_bus_number - 1, axis=0)
    j_11 = numpy.delete(j_11, slack_bus_number - 1, axis=1)
    j_12 = numpy.delete(j_12, slack_bus_number - 1, axis=0)
    j_12 = numpy.delete(j_12, slack_bus_number - 1, axis=1)
    j_21 = numpy.delete(j_21, slack_bus_number - 1, axis=0)
    j_21 = numpy.delete(j_21, slack_bus_number - 1, axis=1)
    j_22 = numpy.delete(j_22, slack_bus_number - 1, axis=0)
    j_22 = numpy.delete(j_22, slack_bus_number - 1, axis=1)

    # Combine matrices together.
    j_1 = numpy.concatenate([j_11, j_12], axis=1)
    j_2 = numpy.concatenate([j_21, j_22], axis=1)
    j = numpy.concatenate([j_1, j_2], axis=0)
    return j


def create_j11_matrix(buses, admittances, s_values):
    """Creates the upper-left Jacobian matrix, J_11.

    Args:
        buses: The state of the system buses.
        admittances: The system admittances.
        s_values: The power mismatch equations.

    Returns:
        The Jacobian matrix J_11.
    """
    j_11 = numpy.zeros((len(buses), len(buses))) * 1j
    for src in buses:
        k = src.number - 1
        v_k = numpy.abs(src.voltage)
        theta_k = numpy.angle(src.voltage)
        q_k = s_values[k].imag

        for dst in buses:
            j = dst.number - 1
            v_j = numpy.abs(dst.voltage)
            theta_j = numpy.angle(dst.voltage)
            b_kj = admittances[k][j].imag

            if j != k:
                j_11[k][j] = -v_k * v_j * b_kj * numpy.cos(theta_k - theta_j)
            else:
                j_11[k][j] = -q_k - v_k ** 2 * b_kj

    return j_11


def create_j12_matrix(buses, admittances, s_values):
    """Creates the upper-left Jacobian matrix, J_12.

    Args:
        buses: The state of the system buses.
        admittances: The system admittances.
        s_values: The power mismatch equations.

    Returns:
        The Jacobian matrix J_12.
    """
    j_12 = numpy.zeros((len(buses), len(buses))) * 1j
    for src in buses:
        k = src.number - 1
        v_k = numpy.abs(src.voltage)
        theta_k = numpy.angle(src.voltage)
        p_k = s_values[k].real

        for dst in buses:
            j = dst.number - 1
            theta_j = numpy.angle(dst.voltage)
            b_kj = admittances[k][j].imag

            if j != k:
                j_12[k][j] = v_k * b_kj * numpy.sin(theta_k - theta_j)
            else:
                j_12[k][j] = p_k / v_k

    return j_12


def create_j21_matrix(buses, admittances, s_values):
    """Creates the upper-left Jacobian matrix, J_21.

    Args:
        buses: The state of the system buses.
        admittances: The system admittances.
        s_values: The power mismatch equations.

    Returns:
        The Jacobian matrix J_21.
    """
    j_21 = numpy.zeros((len(buses), len(buses))) * 1j
    for src in buses:
        k = src.number - 1
        v_k = numpy.abs(src.voltage)
        theta_k = numpy.angle(src.voltage)
        p_k = s_values[k].real

        for dst in buses:
            j = dst.number - 1
            v_j = numpy.abs(dst.voltage)
            theta_j = numpy.angle(dst.voltage)
            b_kj = admittances[k][j].imag

            if j != k:
                j_21[k][j] = -v_k * v_j * b_kj * numpy.sin(theta_k - theta_j)
            else:
                j_21[k][j] = p_k

    return j_21


def create_j22_matrix(buses, admittances, s_values):
    """Creates the upper-left Jacobian matrix, J_22.

    Args:
        buses: The state of the system buses.
        admittances: The system admittances.
        s_values: The power mismatch equations.

    Returns:
        The Jacobian matrix J_22.
    """
    j_22 = numpy.zeros((len(buses), len(buses))) * 1j
    for src in buses:
        k = src.number - 1
        v_k = numpy.abs(src.voltage)
        theta_k = numpy.angle(src.voltage)
        q_k = s_values[k].imag

        for dst in buses:
            j = dst.number - 1
            theta_j = numpy.angle(dst.voltage)
            b_kj = admittances[k][j].imag

            if j != k:
                j_22[k][j] = -v_k * b_kj * numpy.cos(theta_k - theta_j)
            else:
                j_22[k][j] = q_k / v_k - b_kj * v_k

    return j_22
