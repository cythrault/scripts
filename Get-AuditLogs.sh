#!/bin/zsh
nodes=$(ls -1 /ifs/.ifsvar/audit/logs | wc -l)
clustername=$(hostname | cut -f1 -d"-")
today=$(date +"%Y%m%d")
for ((i = 1; i <= $nodes; i++)); do
        logfile="/ifs/AuditLogs/$today-$clustername-$i.log"
        isi_audit_viewer -n $i -t protocol > $logfile
        if [ $(wc -l < $logfile) -lt 2 ]; then 
                rm -f $logfile     
        else
                cp $logfile /ifs/DONNEES/ConservationNumerique/AuditLogs 
        fi  
done