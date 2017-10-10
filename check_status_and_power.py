import sys
import paramiko
import time
import getpass
import os
import re
import socket
import win32com.client as win32


ssh = paramiko.SSHClient()
# Addressing SSH host_key inconvenience
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

path = os.path.dirname(os.path.abspath(__file__))

def read_market_file(path,input_file):
    _file = open("%s\input\%s" % (path,input_file), "r")
    content = _file.readlines()
    content_list = list()
    _file.close()
    for line in content:
        content_list.append(line.split(","))
    return content_list

def invoke_excel():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets('Sheet1')
    ws.Name = 'SFP_test_results'
    return wb,ws,excel

def format_excel(ws):
    rng = ws.Range(ws.Cells(row, 1), ws.Cells(row + 3, 1))
    rng.MergeCells = True
    rng.Value = name
    rng.VerticalAlignment = 2
    rng.HorizontalAlignment = 3
    rng.HorizontalAlignment = 3
    rng.ColumnWidth = 20
    ws.Cells(row, 2).Value = "Intf name"
    ws.Cells(row + 1, 2).Value = "Status"
    ws.Cells(row + 2, 2).Value = "Tx"
    ws.Cells(row + 3, 2).Value = "Rx"

def draw_border(ws,row,column):
    rng = ws.Range(ws.Cells(row, 1), ws.Cells(row + 3, column - 1))
    rng.ColumnWidth = 12
    for border_id in range(7, 10):
        rng.Borders(border_id).LineStyle = 1
        rng.Borders(border_id).Weight = 3
    for border_id in range(10, 13):
        rng.Borders(border_id).LineStyle = 3
        rng.Borders(12).Weight = 1


def close_excel(path,wb,excel,name):
    path = r'%s\output\Cable_testing_results_%s.xls' %(path,name)
    wb.SaveAs(path)
    excel.Application.Quit()


def check_credentials():
    global username
    global password
    try:
        username
    except NameError:
        username = getpass.getuser()
    try:
        password
    except NameError:
        password = getpass.getpass(prompt='Please enter your NT password > ')
    return username, password


class router(object):

    def __init__(self,name):
        self.name = name
        self.remote_shell = None

    def connect(self, ssh, username, password):
        self.ssh = ssh
        try:
            print self.name
            self.ssh.connect(self.name, username = username, password = password)
            time.sleep(1)
            self.remote_shell = self.ssh.invoke_shell()
            time.sleep(1)
            # Disable paging
            self.remote_shell.send("terminal length 0\n")
            time.sleep(1)
            # Clear the buffer on screen
            self.remote_shell.recv(10000)
            print ("SSH session with %s established" % self.name)
        except (socket.error, paramiko.AuthenticationException):
            print ('Connection problem')
            pass

    def check_interfaces(self,int_list):
        global column
        for int_num in int_list:
            ws.Cells(row, column).Value = "Ethernet %s" % int_num.strip()
            interface = intf(int_num.strip(),self.remote_shell)
            interface.check_status()
            interface.check_power()
            column += 1

class intf(object):

    def __init__(self,number,remote_shell):
        self.number = number
        self.remote_shell = remote_shell
        self.status = None
        self.rx = None
        self.tx = None

    def check_status(self):
        if not self.remote_shell: return
        self.remote_shell.send('show interface ethernet %s | i Ethernet.*.is\r\n' % self.number)
        time.sleep(3)
        cli_output = re.sub("\x08", "", self.remote_shell.recv(10000))
        try:
            self.status = ((re.search('Ethernet%s\sis\s(.*)' % self.number, cli_output)).group(1)).strip()
            if self.status != "up": ws.Cells(row + 1, column).Interior.ColorIndex = 3
        except AttributeError:
            self.status = "N/A"
            ws.Cells(row + 1, column).Interior.ColorIndex = 6
        ws.Cells(row + 1, column).Value = self.status

    def check_power(self):
        if not self.remote_shell: return
        if self.status != "up":return
        self.remote_shell.send('show interface ethernet %s transceiver details | i Tx.Power|Rx.Power\r\n' % self.number)
        time.sleep(3)
        cli_output = re.sub("\x08", "", self.remote_shell.recv(10000))
        try:
            self.tx = (re.search('Tx\sPower\s*(.*?)\sdBm', cli_output)).group(1)
            if float(self.tx) < -7:
                ws.Cells(row + 2, column).Interior.ColorIndex = 6
            elif float(self.tx) < -11:
                ws.Cells(row + 2, column).Interior.ColorIndex = 3
        except AttributeError:
            self.tx = "N/A"
            ws.Cells(row + 2, column).Interior.ColorIndex = 6
        ws.Cells(row + 2, column).Value = self.tx
        try:
            self.rx = (re.search('Rx\sPower\s*(.*?)\sdBm', cli_output)).group(1)
            if float(self.rx) < -13: ws.Cells(row + 3, column).Interior.ColorIndex = 3
            elif float(self.rx) < -9: ws.Cells(row + 3, column).Interior.ColorIndex = 6
        except AttributeError:
            self.rx = "N/A"
            ws.Cells(row + 3, column).Interior.ColorIndex = 6
        ws.Cells(row + 3, column).Value = self.rx

if __name__ == "__main__":
    for input_file in os.listdir(path + "\input"):
        if input_file.endswith(".csv") and not input_file.startswith("~"):
            devices_int_list = read_market_file(path, input_file)
            username, password = check_credentials()
            wb, ws, excel = invoke_excel()
            row = 2
            for device in devices_int_list:
                column = 3
                name = device.pop(0).strip()
                format_excel(ws)
                arg = router(name)
                arg.connect(ssh,username,password)
                arg.check_interfaces(device)
                draw_border(ws, row, column)
                row += 4
                ssh.close()
        close_excel(path, wb, excel, input_file)


sys.exit(0)