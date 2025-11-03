#!/bin/bash

# Accept date as input, default to yesterday
input_date=${1:-$(date -d "yesterday" +%Y-%m-%d)}

umask 007

# Get input_date date components
year=$(date -d "$input_date" +%Y)
month=$(date -d "$input_date" +%m)
day=$(date -d "$input_date" +%d)

# SAR binary file for the previous day
sar_file="/var/log/sa/sa$day"

# Set output directory
output_dir="/home/metrics/perf/$year/$month/$day"
mkdir -p "$output_dir"

# Export full-day sar data in CSV format
sadf -T -d -- -u -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_cpu.csv"
sadf -T -d -- -r -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_mem.csv"
sadf -T -d -- -d -p -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_disk.csv"
sadf -T -d -- -n DEV -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_net.csv"
sadf -T -d -- -n NFS -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_nfsclient.csv"
sadf -T -d -- -n NFSD -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_nfsserver.csv"
sadf -T -d -- -S -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_swap.csv"
sadf -T -d -- -W -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_swapping.csv"
sadf -T -d -- -B -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_paging.csv"
sadf -T -d -- -q -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_queue.csv"
sadf -T -d -- -b -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_iotrans.csv"
sadf -T -d -- -H -s 00:00:00 -e 23:59:59 "$sar_file" > "$output_dir/sar_hugepage.csv"

