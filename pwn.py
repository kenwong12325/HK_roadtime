from pwn import *

connection = remote(host, port)

# Example:
r = remote('pwn.hsctf.com', 1234)
# Same as 'nc pwn.hsctf.com 1234'