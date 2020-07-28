import socket
import sys

# Create a TCP/IP socket
sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

# Bind the socket to the port
server_address = ('localhost', 9879)
print('starting up on %s port %s' %(server_address))
sock.bind(server_address)

# Listen for incoming connections
sock.listen(1)

# Wait for a connection
print('waiting for a connection')
connection, client_address = sock.accept()

try:
    print('connection from %s on port %s' %(client_address))
    
    data = connection.recv(5)
    data_bytes = bytes(data)
    print('received 0x%s' %(bytes.hex(data_bytes)))
if data:
    if len(data_bytes) == 5:
        if bytes.hex(data_bytes) == "f000001400":
            print('sending response to client...')
            connection.send(bytes.fromhex("029a"))
        else:
            print('wrong input format: 0x%s' %(bytes.hex(data_bytes)))
    else:
        print('wrong input length: 0x%s' %(bytes.hex(data_bytes)))
else:
    print('null input!')
except (KeyboardInterrupt, SystemExit):
    raise
finally:
# Clean up the connection
    connection.close()
