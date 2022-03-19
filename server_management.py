import os, datetime, time, subprocess, re
import wmi

now = datetime.datetime.now

class Logger:
  def __init__(self):
    self.__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))
    self.logfile = open(
    f'{self.__location__}\\log\\server_management-{now().strftime("%Y-%m-%d")}.log', 'w+')
  def log(self, _type, message):
    log_message = f'[{str(_type).upper()}][{now().strftime("%Y-%m-%d %H:%M:%S")}]:::{message}\n'
    self.logfile.write(log_message)

class Server_Management:
  def __init__(self):
    try:
      self.__location__ = os.path.realpath(
      os.path.join(os.getcwd(), os.path.dirname(__file__)))
      self.logger = Logger()
      self.logger.log("info", "Initiating server management")
    except Exception as e:
      self.logger.log("error", f"{e}")

  def connect(self, server, username, password):
    try:
      self.logger.log("info", f"Creating connection to {str(server)}")
      connection = wmi.WMI(server, user=username, password=password,
      namespace="/root/cimv2")
      self.logger.log("info", f"Connection succeeded on {str(server)}")
      return connection
    except Exception as e:
      self.logger.log("error", f"Not able to connect: {e}")

  def start_service(self, servicename, connection):
    try:
      timeout = 0
      service_state = None
      self.logger.log("info", f"Service {servicename} start process initiated")
      for services in connection.Win32_Service(name=f'{servicename}'):
        services.StartService()
      while service_state != "Running" or timeout != 20:
        self.logger.log("info", f"Service starting...")
        service_state = self.check_service(servicename, connection)
        time.sleep(2)
        timeout += 1
      
      self.logger.log("info", f"{servicename} is now {service_state}")
      return service_state
    except Exception as e:
      self.logger.log("error", f"Not able to start {servicename}: {e}")

  def stop_service(self, servicename, connection):
    try:
      timeout = 0
      service_state = None
      self.logger.log("info", f"Service {servicename} stop process initiated")
      for services in connection.Win32_Service(name=f'{servicename}'):
        services.StopService()
      
      while service_state != "Stopped" or timeout != 20:
        self.logger.log("info", f"Service stopping...")
        service_state = self.check_service(servicename, connection)
        time.sleep(2)
        timeout += 1
      
      self.logger.log("info", f"{servicename} is now {service_state}")
      return service_state
        
    except Exception as e:
      self.logger.log("error", f"Not able to stop {servicename}: {e}")
  
  def check_service(self, servicename, connection):
    try:
      self.logger.log("info", f"Starting to check service {servicename}...")
      wql = f'SELECT * FROM Win32_Service where Name = "{servicename}"'
      for item in connection.query(wql):
        self.logger.log("info", f"Checking service...")
        service_state = str(item.State)
      
      self.logger.log("info", f"{servicename} is now {service_state}")
      return service_state
    except Exception as e:
      self.logger.log("error", f"Not able to check {servicename}: {e}")

  
  def get_cpu_usage(self, connection):
    try:
      self.logger.log("info", f"Starting to check CPU usage")
      wql = f"select * from Win32_PerfRawData_PerfOS_Processor where Name = '_Total'"
      for item in connection.query(wql):
        self.logger.log("info", f"Collecting first CPU sample")
        s1_percent_proc_time = item.PercentProcessorTime
        s1_timestamp = item.Timestamp_Sys100NS
      time.sleep(5)
      for item in connection.query(wql):
        self.logger.log("info", f"Collecting second CPU Sample")
        s2_percent_proc_time = item.PercentProcessorTime
        s2_timestamp = item.Timestamp_Sys100NS
      
      self.logger.log("info", f"Calculating CPU Usage")
      cpu_usage = (1 - ((float(s2_percent_proc_time) - float(s1_percent_proc_time))/(float(s2_timestamp)-float(s1_timestamp))))*100
      self.logger.log("info", f"CPU usage is {cpu_usage}")
      return int(round(cpu_usage, 2))
    except Exception as e:
      self.logger.log("error", f"Not able to get CPU usage: {e}")

  def get_memory_usage(self, connection):
    try:
      self.logger.log("info", f"Starting to get memory usage")
      wql = f"select TotalVisibleMemorySize, FreePhysicalMemory from Win32_OperatingSystem"
      for item in connection.query(wql):
        self.logger.log("info", f"Collecting memory sample")
        memory_usage = 100 - (float(item.FreePhysicalMemory) / float(item.TotalVisibleMemorySize) * 100)
      
      self.logger.log("info", f"Memory usage is {memory_usage}")
      return round(memory_usage, 2)
    except Exception as e:
      self.logger.log("error", f"Not able to get memory usage: {e}")
    
  
  def restart_service(self, servicename, connection):
    try:
      self.logger.log("info", f"Starting to restart service {servicename}")
      service_state = self.stop_service(servicename=servicename, connection=connection)
      service_state = self.start_service(servicename=servicename, connection=connection)
      return service_state
    except Exception as e:
      self.logger.log("error", f"Not able to restart service: {e}")

  def get_boot_time(self, connection):
    try:
      self.logger.log("info", f"Starting to get Last Boot Time")
      for services in connection.win32_operatingsystem():
        boottime = services.LastBootUpTime
        boottime = wmi.to_time(boottime)
        boottime =  datetime.datetime(*boottime[0:6])
      self.logger.log("info", f"Boot Time Pulled")
      
      self.logger.log("info", f"Boot time is {boottime}")
      return boottime
    except Exception as e:
      self.logger.log("error", f"Not able to get Boot Time: {e}")
  
  def get_free_disk_space(self, connection, disk):
    try:
      self.logger.log("info", f"Starting to get free disk space on {disk} drive")
      wql = f"Select Name, FreeSpace, Size from Win32_LogicalDisk where DeviceID = '{str(disk).upper()}:'"
      for item in connection.query(wql):
        free_space = int(item.FreeSpace)/1024/1024/1024
        size = int(item.Size)/1024/1024/1024
        self.logger.log("info", f"Free disk space information acquired")
        
      disk = {
        'Free Space': f"{round(free_space, 2)} GB",
        'Size': f"{round(size, 2)} GB",
        'Percent Free': f"{round(((free_space/size)*100),2)}%"
      }
      self.logger.log("info", f"Disk space is {disk}")
      return disk
    except Exception as e:
      self.logger.log("error", f"Not able to get Free disk space on drive {disk}: {e}")
  
  def get_logged_users(self, connection):
    try:
      users = []
      self.logger.log("info", f"Starting to get logged on users")
      wql = f"Select * from Win32_computersystem"
      for item in connection.query(wql):
        if item.UserName not in users:
          users.append(item.UserName)
          self.logger.log("info", f"User {item.UserName} is logged on")
        self.logger.log("info", f"Logged on users information acquired")

      return users
    except Exception as e:
      self.logger.log("error", f"Not able to get logged on users: {e}")

  def reboot_server(self, servername, credentials):
    try:
      username = credentials['username']
      password = credentials['password']
      script = f"""
      $userName = "{username}"       
      $password = "{password}"
      $secPasswd = ConvertTo-SecureString "$password" -AsPlainText -Force
      $credentials = New-Object System.Management.Automation.PSCredential ("$userName", $secPasswd)
      Restart-Computer -ComputerName {servername} -Credential $credentials -Force
      """
      self.logger.log("info", f"Starting to reboot server {servername}")
      completed = subprocess.run(["powershell", "-Command", script], capture_output=True)
      
      if completed.returncode != 0:
        self.logger.log("error", f"unable to reboot server: {completed.stderr}")
        return completed.stderr
      else:
        self.logger.log("info", f"Reboot command sent successfully!: exited with code {completed.returncode}")
        return f"Command sent successfully!: exited with code {completed.returncode}"
    except Exception as e:
      self.logger.log("error", f"Not able to reboot server: {e}")

  def ping_server(self, servername):
    try:
      self.logger.log("info", f"Starting to ping server {servername}")
      output = os.popen(f'ping {servername}').read()
      self.logger.log("info", f"server ping completed")
      number = re.findall(r'= (\d)', output)
      print(number)
      if number:
        ping_statistics = {
          'Sent': number[0],
          'Received': number[1],
          'Lost': number[2],
          'Minimum': f"{number[3]}ms",
          'Maximum': f"{number[4]}ms",
          'Average': f"{number[5]}ms"
        }
        self.logger.log("info", f"Ping statistics are: {ping_statistics}")
        return ping_statistics
      else:
        return 'not able to ping'
      
    except Exception as e:
      self.logger.log("error", f"Not able to ping server: {e}")

  def traceroute(self, servername):
    try:
      self.logger.log("info", f"Starting to trace route to server {servername}")
      output = os.popen(f'tracert {servername}').read()
      self.logger.log("info", f"server trace route completed")
      print(output)
      
      
      self.logger.log("info", f"traceroute statistics are: {output}")
      return output
    except Exception as e:
      self.logger.log("error", f"Not able to trace route to server: {e}")

  def get_processes(self, connection):
    try:
      processes = []
      self.logger.log("info", f"Starting to get processes")
      wql = f"Select * from Win32_Process"
      for item in connection.query(wql):
        process = {
          'Name': item.Name,
          'Process ID': item.ProcessId,
          }
        processes.append(process)

        self.logger.log("info", f"Processes information acquired: {processes}")

      return processes
    except Exception as e:
      self.logger.log("error", f"Not able to get processes running: {e}")

  def get_hostname(self, connection):
    try:
      self.logger.log("info", f"Starting to get hostname")
      wql = f"SELECT * FROM Win32_ComputerSystem"
      for item in connection.query(wql):
        FQDN = f"{item.Name}.{item.Domain}"
        
      
      self.logger.log("info", f"hostname acquired: {FQDN}")

      return FQDN
    except Exception as e:
      self.logger.log("error", f"Not able to get hostname: {e}")

  def get_recent_events(self, connection):
    try:
      processes = []
      self.logger.log("info", f"Start to get recent events")
      wql = f"Select * Win32_NTLogEvent"
      for item in connection.query(wql):
        print(item)

        self.logger.log("info", f"Events acquired: {item}")

      return item
    except Exception as e:
      self.logger.log("error", f"Not able to get processes running: {e}")