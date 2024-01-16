import math

def calculate_lte_throughput(bandwidth, mimo_layers, mcs, num_rrc_users):
    # Convert bandwidth from MHz to Hz
    bandwidth = bandwidth * 1000000

    # Calculate the number of resource blocks based on bandwidth
    num_resource_blocks = math.floor(bandwidth / 180000)

    # Calculate the maximum number of bits per symbol based on MCS
    if mcs <= 9:
        max_bits_per_symbol = 2
    elif mcs <= 16:
        max_bits_per_symbol = 4
    elif mcs <= 28:
         max_bits_per_symbol =6
    else:
        max_bits_per_symbol = 8

    # Calculate the maximum number of bits per resource element
    max_bits_per_re = max_bits_per_symbol * mimo_layers

    # Calculate the maximum number of bits per resource block
    max_bits_per_rb = max_bits_per_re * 12

    # Calculate the maximum number of bits per subframe
    max_bits_per_sf = max_bits_per_rb * num_resource_blocks

    # Calculate the maximum number of bits per second for one RRC user
    max_bits_per_sec = max_bits_per_sf * 1000 / 10

    # Calculate the maximum number of bits per second for all RRC users
    max_bits_per_sec_total = max_bits_per_sec * num_rrc_users

    # Convert to Mbps and round to two decimal places
    max_throughput = round(max_bits_per_sec_total / 1000000, 2)

    return max_throughput
mmx_tput = calculate_lte_throughput(bandwidth=15, mimo_layers=2, mcs=14, num_rrc_users=480)
print("the maximum throughput for the given configuration is {} Mbps".format(mmx_tput))