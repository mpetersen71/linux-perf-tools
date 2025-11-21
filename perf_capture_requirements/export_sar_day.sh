#!/bin/bash

# Accept date as input, default to yesterday
input_date=${1:-$(date -d "yesterday" +%Y-%m-%d)}

umask 007

# Get input_date date components
year=$(date -d "$input_date" +%Y)
month=$(date -d "$input_date" +%m)
day=$(date -d "$input_date" +%d)

# SAR binary file for the previous day, also check for sysstat instead of sa
sar_file="/var/log/sa/sa$day"
if [ -d "/var/log/sysstat" ]; then
	sar_file="/var/log/sysstat/sa$day"
fi

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

# Collect System Info

DATE=$(date +"%Y-%m-%d %H:%M:%S")

# === General System Info ===
HOSTNAME=$(hostname)
OS=$(grep PRETTY_NAME /etc/os-release | cut -d= -f2 | tr -d '"')
KERNEL=$(uname -r)
ARCH=$(uname -m)
UPTIME_HOURS=$(awk '{print int($1/3600)}' /proc/uptime)
UPTIME_PRETTY=$(uptime -p)
LAST_BOOT=$(who -b | awk '{print $3, $4}')
CPUS=$(nproc)
CPU_MODEL=$(grep -m1 "model name" /proc/cpuinfo | cut -d: -f2 | xargs)
MEM_TOTAL_KB=$(grep MemTotal /proc/meminfo | awk '{print $2}')
MEM_TOTAL_MB=$((MEM_TOTAL_KB / 1024))

cat <<EOF > "$output_dir/sysinfo.xml"
<system_report>
  <date>$DATE</date>
  <hostname>$HOSTNAME</hostname>
  <os>$OS</os>
  <kernel>$KERNEL</kernel>
  <architecture>$ARCH</architecture>
  <uptime_hours>$UPTIME_HOURS</uptime_hours>
  <uptime_pretty>$UPTIME_PRETTY</uptime_pretty>
  <last_boot>$LAST_BOOT</last_boot>
  <cpus>$CPUS</cpus>
  <cpu_model>$CPU_MODEL</cpu_model>
  <memory_total_mb>$MEM_TOTAL_MB</memory_total_mb>
</system_report>
EOF

# === Disk Usage ===
DISK_USAGE=$(df -h --output=source,size,used,avail,pcent,target | tail -n +2)

cat <<EOF > "$output_dir/disk_usage.xml"
<disk_usage_report>
  <date>$DATE</date>
  <disk_usage>
$(echo "$DISK_USAGE" | awk '{printf "    <disk><device>%s</device><size>%s</size><used>%s</used><available>%s</available><percent>%s</percent><mount>%s</mount></disk>\n", $1,$2,$3,$4,$5,$6}')
  </disk_usage>
</disk_usage_report>
EOF


# === Running Services ===
RUNNING_SERVICES=$(systemctl list-units --type=service --state=running --no-legend | awk '{print $1}')

cat <<EOF > "$output_dir/services.xml"
<running_services_report>
  <date>$DATE</date>
  <services>
$(echo "$RUNNING_SERVICES" | awk '{printf "    <service>%s</service>\n", $1}')
  </services>
</running_services_report>
EOF


# === Running Processes ===
PROCESSES=$(ps -eo pid,user,%cpu,%mem,cmd --no-headers --sort=-%cpu)

cat <<EOF > "$output_dir/processes.xml"
<running_processes_report>
  <date>$DATE</date>
  <processes>
$(echo "$PROCESSES" | awk '{
    pid=$1; user=$2; cpu=$3; mem=$4;
    $1=$2=$3=$4=""; cmd=substr($0,5);
    gsub("&","&amp;",cmd); gsub("<","&lt;",cmd); gsub(">","&gt;",cmd);
    printf "    <process><pid>%s</pid><user>%s</user><cpu>%s</cpu><mem>%s</mem><command>%s</command></process>\n", pid, user, cpu, mem, cmd
}')
  </processes>
</running_processes_report>
EOF


# === Listening Ports ===
LISTENING_PORTS=$(ss -tulpnH | awk '{print $1,$5,$7}')

cat <<EOF > "$output_dir/listening_ports.xml"
<listening_ports_report>
  <date>$DATE</date>
  <listening_ports>
$(echo "$LISTENING_PORTS" | awk '{
    split($2,addr,":");
    process=$3;
    gsub("\"","",process);
    printf "    <port><protocol>%s</protocol><ip>%s</ip><port_number>%s</port_number><process>%s</process></port>\n", $1, addr[1], addr[2], process
}')
  </listening_ports>
</listening_ports_report>
EOF




