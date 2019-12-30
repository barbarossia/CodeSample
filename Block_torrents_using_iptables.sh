# Block torrents using iptables

# Log Torrent

iptables -N LOGDROP > /dev/null 2> /dev/null
iptables -F LOGDROP
iptables -A LOGDROP -j LOG --log-prefix "LOGDROP "
iptables -A LOGDROP -j DROP


# Block Torrent

ip=192.168.1.13

iptables -A FORWARD -s $ip -m string --algo bm --string "BitTorrent" -j LOGDROP
iptables -A FORWARD -s $ip -m string --algo bm --string "BitTorrent protocol" -j LOGDROP
iptables -A FORWARD -s $ip -m string --algo bm --string "peer_id=" -j LOGDROP
iptables -A FORWARD -s $ip -m string --algo bm --string ".torrent" -j LOGDROP
iptables -A FORWARD -s $ip -m string --algo bm --string "announce.php?passkey=" -j LOGDROP
iptables -A FORWARD -s $ip -m string --algo bm --string "torrent" -j LOGDROP
iptables -A FORWARD -s $ip -m string --algo bm --string "announce" -j LOGDROP
iptables -A FORWARD -s $ip -m string --algo bm --string "info_hash" -j LOGDROP

# Block DHT keyword

iptables -A FORWARD -s $ip -m string --string "get_peers" --algo bm -j LOGDROP
iptables -A FORWARD -s $ip -m string --string "announce_peer" --algo bm -j LOGDROP
iptables -A FORWARD -s $ip -m string --string "find_node" --algo bm -j LOGDROP