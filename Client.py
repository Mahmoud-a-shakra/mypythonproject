import socket
import sys

def validIPAddress(IP):
    def isIPv4(s):
        try: return str(int(s)) == s and 0 <= int(s) <= 255
        except: return False
    def isIPv6(s):
        if len(s) > 4:
            return False
        try : return int(s, 16) >= 0 and s[0] != '-'
        except:
            return False
    if IP.count(".") == 3 and all(isIPv4(i) for i in IP.split(".")):
        return " Valid IP address-IPv4"
    if IP.count(":") == 7 and all(isIPv6(i) for i in IP.split(":")):
        return "Valid IP address-IPv6"
    return "Not valid"Ip_value=input('Enter your IP address:')print(validIPAddress(Ip_value))port_value=input('Enter your Port Value (optinal):')
if port_value=='':
    port_value=6653
print ("your port value is: ")
print(port_value)
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
print ("your port value is: ")
print(port_value)


# Create a TCP/IP socket
sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

# Connect the socket to the port where the server is listening
server_address = ('localhost', 10000)
print >>sys.stderr, 'connecting to %s port %s' % server_address
sock.connect(server_address)

try:
    
    # Send data
    message = 'This is the message.  It will be repeated.'
    print >>sys.stderr, 'sending "%s"' % message
    sock.sendall(message)

    # Look for the response
    amount_received = 0
    amount_expected = len(message)
    
    while amount_received < amount_expected:
        data = sock.recv(16)
        amount_received += len(data)
        print >>sys.stderr, 'received "%s"' % data

finally:
    print >>sys.stderr, 'closing socket'
    sock.close()
