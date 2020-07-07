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
    return "Not valid"
Ip_value=input('Enter your IP address:')
print(validIPAddress(Ip_value))
port_value=input('Enter your Port Value (optinal):')
if port_value=='':
    port_value=6653
print ("your port value is: ")
print(port_value)
