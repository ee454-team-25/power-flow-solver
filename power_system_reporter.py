"""A module containing power system reports."""

import numpy
import power_flow_solver
import tabulate

TABULATE_FLOAT_FMT = '.4f'


def bus_voltage_report(system):
    """Reports the voltage at each bus.

    Args:
        system: The power system being analyzed.
    """
    headers = ['Bus', 'Voltage (V)', 'Phase (deg)']
    table = [[bus.number, numpy.abs(bus.voltage), numpy.rad2deg(numpy.angle(bus.voltage))] for bus in system.buses]
    return tabulate.tabulate(table, headers=headers, floatfmt=TABULATE_FLOAT_FMT)


def line_power_report(system, power_base):
    """Reports the active, reactive, and apparent power for each transmission line.

    TODO(kjiwa): Report whether a line has exceeded its power rating.

    Args:
        system: The power system being analyzed.
        power_base: The power base in MVA.
    """
    headers = ['Line', 'Active Power (MW)', 'Reactive Power (Mvar)', 'Apparent Power (MVA)']
    table = []
    for line in system.lines:
        src = system.buses[line.source - 1]
        dst = system.buses[line.destination - 1]

        v = dst.voltage - src.voltage
        complex_power = power_base * numpy.abs(v) ** 2 / numpy.conj(line.distributed_impedance)
        line_name = '{}-{}'.format(src.number, dst.number)
        table.append([line_name, complex_power.real, complex_power.imag, numpy.abs(complex_power)])

    return tabulate.tabulate(table, headers=headers, floatfmt=TABULATE_FLOAT_FMT)


def largest_power_mismatch_report(estimates, power_base, iteration):
    """Reports the largest active and reactive power mismatches for a set of estimates.

    Args:
        estimates: A set of bus power estimates.
        power_base: The power base in MVA.
        iteration: The current iteration of the power flow solver.
    """
    max_p_error = 0
    max_p_error_bus = None

    max_q_error = 0
    max_q_error_bus = None

    for estimate in estimates.values():
        if estimate.active_power_error > max_p_error:
            max_p_error = estimate.active_power_error
            max_p_error_bus = estimate.bus.number

        if estimate.reactive_power_error > max_q_error:
            max_q_error = estimate.reactive_power_error
            max_q_error_bus = estimate.bus.number

    max_p_error *= power_base
    max_q_error *= power_base

    s = 'Iteration {}'.format(iteration)
    if max_p_error:
        s += '\n  Largest active power mismatch:   {:.4f} MW (Bus {})'.format(max_p_error, max_p_error_bus)

    if max_q_error:
        s += '\n  Largest reactive power mismatch: {:.4f} Mvar (Bus {})'.format(max_q_error, max_q_error_bus)

    return s


def power_injection_report(system, estimates, power_base):
    """Reports the active and reactive power injection from each generator and synchronous consider.

    Args:
        estimates: A dict mapping bus numbers to estimates of its active power injection.
        power_base: The power base in MVA.
    """
    headers = ['Bus', 'Active Power Injection (MW)', 'Reactive Power Injection (Mvar)']
    table = []
    for estimate in estimates.values():
        if estimate.bus_type != power_flow_solver.BusType.PV:
            continue

        q_injected = -(estimate.reactive_power - estimate.bus.reactive_power_consumed) * power_base
        table.append([estimate.bus.number, estimate.bus.active_power_injected * power_base, q_injected])

    return tabulate.tabulate(table, headers=headers, floatfmt=TABULATE_FLOAT_FMT)
