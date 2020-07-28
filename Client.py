import socket
import sys

def validIPAddress(IP):
    def isIPv4(s):
        try: return str(int(s)) == s and 0 <= int(s) <= 255
        except: return False
    
    if IP.count(".") == 3 and all(isIPv4(i) for i in IP.split(".")):
        return " Valid IP address-IPv4"
    
    return "Not valid"
Ip_value=input('Enter your IP address:')
print(validIPAddress(Ip_value))
port_value=input('Enter your Port Value (optinal):')
if port_value=='':
    port_value=6653
print ("your port value is: %s" %port_value)


# Create a TCP/IP socket
sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

# Connect the socket to the port where the server is listening
server_address = ('localhost', 9879)
print('connecting to %s port %s' %(server_address))
sock.connect(server_address)

try:
    # Send data
    hex_str = 'f000001400'
    hex_bytes = bytes.fromhex(hex_str)
    print('sending 0x%s' %(bytes.hex(hex_bytes)))
    sock.send(hex_bytes)
    
    data = sock.recv(2)
    data_bytes = bytes(data)
    print('received in bytes: 0x%s' %(bytes.hex(data_bytes)))
    print('received in decimal (big-endian): %s' %(int.from_bytes(data_bytes, "big")))
    print('received in decimal (little-endian): %s' %(int.from_bytes(data_bytes, "little")))

finally:
    print('closing socket')
    sock.close()
