#!/bin/bash

myip=$(ipconfig getifaddr en0)
#op=$1
#echo $op
if [ "$1" == "add" ];then
	/usr/bin/ssh root@192.168.1.1 'sh -s' < /Users/barbarossia/add_to_gfwlist.sh $myip
	echo "add ip to gfwlist now"
elif [ "$1" == "remove" ];then
	/usr/bin/ssh root@192.168.1.1 'sh -s' < /Users/barbarossia/remove_to_gfwlist.sh $myip
	echo "remove ip to gfwlist now"
else
	echo "error input"
fi
