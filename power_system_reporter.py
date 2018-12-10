"""A module containing power system reports."""

import numpy
import power_flow_solver
import tabulate

TABULATE_FLOAT_FMT = '.4f'


def bus_voltage_report(system, min_operating_voltage, max_operating_voltage):
    """Reports the voltage at each bus and whether the bus is outside operating limits.

    Args:
        system: The power system being analyzed.
        min_operating_voltage: The minimum operating voltage.
        max_operating_voltage: The maximum operating voltage.
    """
    headers = ['Bus', 'Voltage (V)', 'Phase (deg)', 'Outside Operating Limits']
    table = []
    for bus in system.buses:
        voltage = numpy.abs(bus.voltage)
        outside_limits = 'Yes' if voltage < min_operating_voltage or voltage > max_operating_voltage else 'No'
        table.append([bus.number, numpy.abs(bus.voltage), numpy.rad2deg(numpy.angle(bus.voltage)), outside_limits])

    return tabulate.tabulate(table, headers=headers, floatfmt=TABULATE_FLOAT_FMT)


def line_power_report(system, power_base):
    """Reports the active, reactive, and apparent power for each transmission line.

    Args:
        system: The power system being analyzed.
        power_base: The power base in MVA.
    """
    headers = ['Line', 'Sending Power (MW)', 'Sending Power (Mvar)', 'Sending Power (MVA)', 'Receiving Power (MW)',
               'Receiving Power (Mvar)', 'Receiving Power (MVA)', 'Exceeds Rating']
    table = []
    for line in system.lines:
        src = system.buses[line.source - 1]
        dst = system.buses[line.destination - 1]

        i_src = (src.voltage - dst.voltage) / line.distributed_impedance + src.voltage * line.shunt_admittance
        i_dst = (dst.voltage - src.voltage) / line.distributed_impedance + dst.voltage * line.shunt_admittance
        s_src = power_base * src.voltage * numpy.conj(i_src)
        s_dst = power_base * dst.voltage * numpy.conj(i_dst)

        line_name = '{}-{}'.format(src.number, dst.number)
        exceeds_rating = 'Yes' if line.max_power and max(s_src, s_dst) > line.max_power else 'No'
        table.append([line_name, s_src.real, s_src.imag, numpy.abs(s_src), s_dst.real, s_dst.imag, numpy.abs(s_dst),
                      exceeds_rating])

    return tabulate.tabulate(table, headers=headers, floatfmt=TABULATE_FLOAT_FMT)


def largest_power_mismatch_report(estimates, power_base, iteration):
    """Reports the largest active and reactive power mismatches for a set of estimates.

    Args:
        estimates: A set of bus power estimates.
        power_base: The power base in MVA.
        iteration: The current iteration of the power flow solver.
    """
    # Find the maximum active power error.
    pv_pq_estimates = [i for i in estimates.values() if i.bus_type != power_flow_solver.BusType.SWING]
    max_p_error_estimate = max(pv_pq_estimates, key=lambda x: numpy.abs(x.active_power_error))
    max_p_error = max_p_error_estimate.active_power_error * power_base

    # Find the maximum reactive power error.
    pq_estimates = [i for i in estimates.values() if i.bus_type == power_flow_solver.BusType.PQ]
    max_q_error_estimate = max(pq_estimates, key=lambda x: numpy.abs(x.reactive_power_error))
    max_q_error = max_q_error_estimate.reactive_power_error * power_base

    return '''Iteration {}
  Largest active power mismatch:   {:8.4f} MW   (Bus {})
  Largest reactive power mismatch: {:8.4f} Mvar (Bus {})'''.format(
        iteration, max_p_error, max_p_error_estimate.bus.number, max_q_error, max_q_error_estimate.bus.number)


def power_injection_report(estimates, power_base, max_active_power_error, max_reactive_power_error):
    """Reports the active and reactive power injection from each generator and synchronous condenser.

    Args:
        estimates: A dict mapping bus numbers to estimates of its active power injection.
        power_base: The power base in MVA.
        max_active_power_error: The maximum allowed active power error in MW.
        max_reactive_power_error: The maximum allowed reactive power error in Mvar.
    """
    headers = ['Bus', 'Power Injection (MW)', 'Power Injection (Mvar)']
    table = []
    for estimate in estimates.values():
        if estimate.bus_type == power_flow_solver.BusType.PQ:
            continue

        p_injected = (estimate.active_power + estimate.bus.active_power_consumed) * power_base
        q_injected = (estimate.reactive_power + estimate.bus.reactive_power_consumed) * power_base
        table.append([estimate.bus.number, p_injected, q_injected])

    floatfmt_p = '.{}f'.format(int(numpy.ceil(-numpy.log10(max_active_power_error))))
    floatfmt_q = '.{}f'.format(int(numpy.ceil(-numpy.log10(max_reactive_power_error))))
    return tabulate.tabulate(table, headers=headers, floatfmt=(None, floatfmt_p, floatfmt_q))
