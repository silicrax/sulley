#!c:\\python\\python.exe

import os
import sys
import time
import getopt
from vboxapi import VirtualBoxManager

try:
    from win32api import GetShortPathName
    from win34com.shell import shell
except:
    if os.name == "nt":
        print "[!] Failed to import win32api/win32com modules, please install these! Bailing..."
        sys.exit(1)


from sulley import pedrpc

PORT  = 26003
ERR   = lambda msg: sys.stderr.write("ERR> " + msg + "\n") or sys.exit(1)
USAGE = "USAGE: vboxcontrol.py"                                                                  \
        "\n    <-v|--vbox NAME>       set the name of the vbox to control"                       \
        "\n    <-m|--vbm FILENAME>    path to vboxmanage executable"                             \
        "\n    [-s|--snapshot NAME]   set the snapshot name"                                     \
        "\n    [-l|--log_level LEVEL] log level (default 1), increase for more verbosity"        \
        "\n    [-c|--restorecurrent]  Restore current snapshot, can be used instead of snapshot" \
        "\n    [--port PORT]          TCP port to bind this agent to"


########################################################################################################################
class vmcontrol_pedrpc_server (pedrpc.server):
    def __init__ (self, host, port, vboxmanage, vbox, snap_name=None, log_level=1, restorecurrent=False):
        '''
        @type  host:            String
        @param host:            Hostname or IP address to bind server to
        @type  port:            Integer
        @param port:            Port to bind server to
        @type  vboxmanage:      String
        @param vboxmanage:      Path to VirtualBox vboxmange executable
        @type  vbox:            String
        @param vbox:            Name of VirtualBox vbox
        @type  snap_name:       String
        @param snap_name:       (Optional, def=None) Snapshot name to revert to on restart
        @type  log_level:       Integer
        @param log_level:       (Optional, def=1) Log output level, increase for more verbosity
        @type  restorecurrent:  Boolean
        @param restorecurrent:  (Option, def=False) Revert to current snapshot
        '''

        # initialize the PED-RPC server.
        pedrpc.server.__init__(self, host, port)

        self.host           = host
        self.port           = port

        self.restorecurrent = restorecurrent

        # if we're on windows, get the DOS path names
        if os.name == "nt":
            self.vboxmanage = GetShortPathName(r"%s" % vboxmanage)
            self.vbox   = GetShortPathName(r"%s" % vbox)
        else:
            self.vboxmanage = vboxmanage
            self.vbox   = vbox

        self.snap_name   = snap_name
        self.log_level   = log_level
        self.restorecurrent = restorecurrent

        self.log("VboxManage PED-RPC server initialized:")
        self.log("\t vboxmanage: %s" % self.vboxmanage)
        self.log("\t vbox:       %s" % self.vbox)
        if self.restorecurrent:
            self.log("\t snap name:  current")
        else:
            self.log("\t snap name:  %s" % self.snap_name)
        self.log("\t log level:  %d" % self.log_level)
        self.log("Awaiting requests...")

        virtboxManager = VirtualBoxManager(None, None)
        virtbox = virtboxManager.vbox
        virtboxMgr = virtboxManager.mgr
        constants = virtboxManager.constants
        virtboxMachine = virtbox.findMachine(self.vbox)

	self.virtboxMachine = virtboxMachine


    def alive (self):
        '''
        Returns True. Useful for PED-RPC clients who want to see if the PED-RPC connection is still alive.
        '''

        return True


    def log (self, msg="", level=1):
        '''
        If the supplied message falls under the current log level, print the specified message to screen.

        @type  msg: String
        @param msg: Message to log
        '''

        if self.log_level >= level:
            print "[%s] %s" % (time.strftime("%I:%M.%S"), msg)


    def vbcommand (self, command):
        '''
        Execute the specified command, keep trying in the event of a failure.

        @type  command: String
        @param command: vboxmanage command to execute
        '''

        while 1:
            self.log("executing: %s" % command, 5)

            pipe = os.popen(command)
            out  = pipe.readlines()

            try:
                pipe.close()
            except IOError:
                self.log("IOError trying to close pipe")

            if not out:
                break
            elif not out[0].lower().startswith("close failed"):
                break

            self.log("failed executing command '%s' (%s). will try again." % (command, out))
            time.sleep(1)

        return "".join(out)


    ###
    ### VMRUN COMMAND WRAPPERS
    ###


    def revert_to_snapshot (self, snap_name=None):
	if self.restorecurrent:
            self.log("reverting to snapshot: %s" % snap_name, 2)
            command = self.vboxmanage + " snapshot " + self.vbox + " restorecurrent"
            print command
            return self.vbcommand(command)
        else:
            if not snap_name:
                snap_name = self.snap_name
            self.log("reverting to snapshot: %s" % snap_name, 2)
            command = self.vboxmanage + " snapshot " + self.vbox + " restore " + snap_name
            print command
            return self.vbcommand(command)


    def start (self):
        self.log("starting image", 2)

        command = self.vboxmanage + " startvm " + self.vbox
        return self.vbcommand(command)


    def stop (self):
        self.log("stopping image", 2)

        command = self.vboxmanage + " controlvm " + self.vbox + " poweroff"
        return self.vbcommand(command)


    ###
    ### EXTENDED COMMANDS
    ###


    def restart_target (self):
        self.log("restarting virtual machine...")

        # revert to the specified snapshot and start the image.
        self.stop()
        self.wait_stop()
        self.revert_to_snapshot()
        self.wait_restore()
        self.start()
        self.wait_start()


    def is_target_ready (self):
        # sometimes vboxmanage reports that the VM is up while it's still reverting.
        time.sleep(10)
        return self.virtboxMachine.state


    def wait_stop (self):
        self.log("waiting for vbox %s to stop" % self.vbox)
        while 1:
            if self.is_target_ready() == 1:
                break


    def wait_restore (self):
        self.log("waiting for vbox %s to be restored" % self.vbox)
        while 1:
            if self.is_target_ready() == 2:
                break


    def wait_start (self):
        self.log("waiting for vbox %s to come up" % self.vbox)
        while 1:
            if self.is_target_ready() == 5:
                break



########################################################################################################################
if __name__ == "__main__":
    # parse command line options.
    try:
        opts, args = getopt.getopt(sys.argv[1:], "v:m:s:l:c", ["vbox=", "vboxmanage=", "snapshot=", "log_level=", "restorecurrent", "port="])
    except getopt.GetoptError:
        print "error here"
        ERR(USAGE)

    if sys.platform.startswith("linux"):
        vboxmanage  = "/usr/bin/vboxmanage"
    else:
        vboxmanage = r"C:\progra~1\oracle\virtua~1\vboxmanage.exe"

    vbox           = None
    snap_name      = None
    log_level      = 1
    restorecurrent = False

    for opt, arg in opts:
        if opt in ("-v", "--vbox"):           vbox           = arg
        if opt in ("-m", "--vbm"):            vboxmanage     = arg
        if opt in ("-s", "--snapshot"):       snap_name      = arg
        if opt in ("-l", "--log_level"):      log_level      = int(arg)
        if opt in ("-c", "--restorecurrent"): restorecurrent = True
        if opt in ("--port"):                 PORT           = int(arg)

    # OS check

    if not vbox and (not restorecurrent or not snap_name):
        print "error there"
        ERR(USAGE)

    servlet = vmcontrol_pedrpc_server("0.0.0.0", PORT, vboxmanage, vbox, snap_name, log_level, restorecurrent)
    servlet.serve_forever()
