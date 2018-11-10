import enum
import namedlist
import numpy
import operator


class BusType(enum.Enum):
    """An enumeration of bus types."""
    UNKNOWN = 0
    LOAD = 1
    GENERATOR = 2
    SLACK = 3


# An object containing load bus data.
#
# Attributes:
#   type: The bus type.
#   number: The bus number.
#   p: The real power consumed (MW).
#   q: The reactive power consumed (Mvar).
#   voltage: The bus voltage (pu).
Bus = namedlist.namedlist('Bus', ['type', 'number', 'p', 'q', 'voltage'])


def count_rows(ws):
    """Counts the number of non-empty, contiguous rows in the worksheet.

    This method is necessary because openpyxl.worksheet.worksheet.max_row does not always correctly detect empty cells.

    Args:
        ws: The worksheet to read.

    Returns:
        The number of non-empty, contiguous rows.
    """
    n = 0
    for row in ws.iter_rows():
        if not row[0].value:
            break

        n += 1

    return n


def read_admittance_matrix(ws, num_buses):
    """Creates an admittance matrix from an input worksheet.

    The worksheet is expected to have a header row. Each subsequent row is expected to contain the following columns:

        1. Source bus
        2. Destination bus
        3. Total resistance (pu)
        4. Total reactance (pu)
        5. Total susceptance (pu)
        6. Maximum apparent power (MVA)

    Furthermore, it is expected line data is not duplicated and that the source and destination buses are different.

    Args:
        ws: The worksheet containing line data.
        num_buses: The number of buses in the system.

    Returns:
        A square matrix containing line-to-line admittances.
    """
    admittances = numpy.zeros((num_buses, num_buses)) * 1j  # multiply by 1j to make values complex

    # Compute matrix entries.
    for row in ws.iter_rows(row_offset=1, max_row=num_buses):
        # Subtract 1 since arrays start at 0, but buses start at 1.
        src = row[0].value - 1
        dst = row[1].value - 1

        rtotal = row[2].value
        xtotal = row[3].value
        btotal = row[4].value
        # fmax = row[5].value

        # Series and parallel impedances.
        y_series = 1 / (rtotal + 1j * xtotal)
        y_parallel = 1j * btotal / 2

        # Subtract series admittances from off-diagonal entries.
        admittances[src][dst] -= y_series
        admittances[dst][src] -= y_series

        # Add series and parallel admittances to diagonal entries.
        admittances[src][src] += y_series + y_parallel
        admittances[dst][dst] += y_series + y_parallel

    return admittances


def read_bus_data(ws, slack_bus_num, initial_voltage, power_base_mva):
    """Reads bus data from an input worksheet.

    The input data is expected to have the following columns:

        1. Bus number
        2. Real power consumed (MW)
        3. Reactive power consumed (Mvar)
        4. Real power delivered (MW)
        5. Voltage (pu)

    The slack bus has only its reference voltage set. The real and reactive power are set to balance the rest of the
    system, and power generation is zero.

    A load (PQ) bus is a bus that has real and reactive power consumption. For these buses, power generation is zero.

    A generator (PV) bus is a bus that has real power generation and a terminal voltage. For these buses, power
    consumption is zero.

    Args:
        ws: The worksheet containing bus data.
        slack_bus_num: The slack bus number.
        initial_voltage: The initial reference voltage.
        power_base_mva: The apparent power base in MVA.

    Returns:
        An array containing load bus data.
    """
    p_total = 0
    q_total = 0

    buses = []
    for i, row in enumerate(ws.iter_rows(row_offset=1, max_row=count_rows(ws))):
        bus_num = row[0].value
        p_load = row[1].value or 0
        q_load = row[2].value or 0
        p_gen = row[3].value or 0
        voltage = row[4].value or initial_voltage

        # Determine bus type.
        type = BusType.UNKNOWN
        if p_load != 0 and q_load != 0 and p_gen == 0:
            type = BusType.LOAD
        elif p_load == 0 and q_load == 0 and p_gen != 0:
            type = BusType.GENERATOR

        # Convert quantities to per-unit.
        p = (p_load - p_gen) / power_base_mva
        q = q_load / power_base_mva

        p_total += p
        q_total += q

        buses.append(Bus(type, bus_num, p, q, voltage))

    # Balance system power in the slack bus.
    buses = sorted(buses, key=operator.attrgetter('number'))
    buses[slack_bus_num - 1].type = BusType.SLACK
    buses[slack_bus_num - 1].p = -p_total
    buses[slack_bus_num - 1].q = -q_total
    return buses
