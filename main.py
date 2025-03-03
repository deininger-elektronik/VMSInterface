
# Weingarten (Baden), Germany FEB. 2025

# **Lead Developer:** [github.com/emsysol](https://github.com/emsysol)  
# Parts of this code were generated using **GitHub Copilot AI** under GitHub 
# user account: **[github.com/batchdeploy](https://github.com/batchdeploy)**

# This project is released under the GNU General Public License v3.0.

# You are free to use, modify, and distribute this software under the terms of
# the GNU GPL v3, either version 3 or (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.

 
# Disclaimer:  
# THIS SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE, AND NON-INFRINGEMENT.  
# IN NO EVENT SHALL THE AUTHORS, COPYRIGHT HOLDERS, OR CONTRIBUTORS BE LIABLE
# FOR ANY CLAIM, DAMAGES, OR OTHER LIABILITY, WHETHER IN AN ACTION OF
# CONTRACT, TORT, OR OTHERWISE, ARISING FROM, OUT OF, OR IN CONNECTION WITH
# THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.  
  
#  **USE THIS SOFTWARE AT YOUR OWN RISK.** No guarantees are made regarding
#  functionality, security, or performance.  

# Parts of this code were generated using **GitHub Copilot AI** under the
# GitHub user account:
# [github.com/batchdeploy](https://github.com/batchdeploy).

# By using this software, you acknowledge that portions were generated with
# the assistance of **GitHub Copilot AI** and accept the associated terms and
# conditions.

import serial
#import json
import time
import win32gui
import win32con
import win32api
import win32process
import logging
import logging.handlers
import os
import yaml
import psutil

from keycombinations_dict import keycombinations    # keycombinations.py
from keycodes_dict import keycodes                  # keycodes.py

def enable_logging():
    try:
        log_directory = os.path.join(os.environ['LOCALAPPDATA'], 'VMSInterface', 'Logs')
        if not os.path.exists(log_directory):
            os.makedirs(log_directory)
        log_file_path = os.path.join(log_directory, 'application.log')

        # Create a rotating file handler
        rotating_handler = logging.handlers.RotatingFileHandler(
            log_file_path, maxBytes=5*1024*1024, backupCount=5, encoding='utf-8'
        )
        rotating_handler.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%d.%m.%y %H:%M:%S')
        formatter.converter = time.localtime
        rotating_handler.setFormatter(formatter)

        # Create Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)
        console_handler.setFormatter(formatter)

        # Configure the root logger
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.DEBUG)
        root_logger.handlers = []  # Clear existing handlers
        root_logger.addHandler(rotating_handler)
        root_logger.addHandler(console_handler)

    except Exception as e:
        logging.debug('Failed to enable logging: ' + str(e))

def load_config():
    try:
        # Load the config.yaml file
        app_name = 'VMSInterface'
        config_file_name = 'config.yaml'
        program_data_path = os.environ.get('PROGRAMDATA', 'C:\\ProgramData')
        config_directory = os.path.join(program_data_path, app_name)
        config_file_path = os.path.join(config_directory, config_file_name)
        if not os.path.exists(config_directory):
            os.makedirs(config_directory)
            logging.error('Config directory does not exist, created directory: ' + config_directory)
        
        if not os.path.exists(config_file_path):
            default_config = {
                "COM_PORT": "COM9",
                "BAUD_RATE": 9600,
                "BYTESIZE": 8,
                "PARITY": "EVEN",
                "STOPBITS": 1,
                "XONXOFF": "False",
                "RTSCTS": "False",
                "DSRDTR": "False",
                "WINDOW_NAME": "Client",
                "LOG_LEVEL": "INFO",
                "WAIT_AFTER_KEY": 1,
                "CHECK_FOR_WINDOW_ON_STARTUP": "True",
                "SET_TO_FOREGROUND_ENABLE": "True",
                "SET_TO_FOREGROUND_PAUSE":30,
                "SET_TO_FOREGROUND_STRICT": "False",
                "MONITOR_COUNT": 1
            }
            with open(config_file_path, 'w') as config_file:
                yaml.dump(default_config, config_file)
                logging.error('Config file does not exist, created file: ' + config_file_path)
        
        with open(config_file_path) as config_file:
            config = yaml.safe_load(config_file)

        # Convert baudrate, bytesize, and stopbits to integers
        config['BAUD_RATE'] = int(config['BAUD_RATE'])
        config['BYTESIZE'] = int(config['BYTESIZE'])
        config['STOPBITS'] = int(config['STOPBITS'])

        # Convert xonxoff, rtscts, and dsrdtr to lowercase and check against true
        config['XONXOFF'] = config['XONXOFF'].lower() == 'true'
        config['RTSCTS'] = config['RTSCTS'].lower() == 'true'
        config['DSRDTR'] = config['DSRDTR'].lower() == 'true'

        # Map the parity entry from EVEN, ODD, NONE to the syntax expected by the packages
        if config['PARITY'] == 'EVEN':
            config['PARITY'] = serial.PARITY_EVEN
        elif config['PARITY'] == 'ODD':
            config['PARITY'] = serial.PARITY_ODD
        else:
            config['PARITY'] = serial.PARITY_NONE

        # Add validation for config['LOG_LEVEL'] to check if it is a valid log level
        if config['LOG_LEVEL'].upper() not in ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']:
            config['LOG_LEVEL'] = 'INFO'  # Set default log level to INFO if invalid log level is specified
            logging.info('Log Level set in config is invalid, setting to default value: INFO')

        # Convert MONITOR_COUNT to integer
        config['MONITOR_COUNT'] = int(config['MONITOR_COUNT'])
        # Convert WAIT_AFTER_KEY to float
        config['WAIT_AFTER_KEY'] = float(config['WAIT_AFTER_KEY'])
        # Convert MONITOR_COUNT to integer
        config['SET_TO_FOREGROUND_PAUSE'] = int(config['SET_TO_FOREGROUND_PAUSE'])

        config['SET_TO_FOREGROUND_STRICT'] = config['SET_TO_FOREGROUND_STRICT'].lower() == 'true'

        config['SET_TO_FOREGROUND_ENABLE'] = config['SET_TO_FOREGROUND_ENABLE'].lower() == 'true'

        config['CHECK_FOR_WINDOW_ON_STARTUP'] = config['CHECK_FOR_WINDOW_ON_STARTUP'].lower() == 'true'

        logging.getLogger().setLevel(config['LOG_LEVEL'].upper()) # change log level to the one specified in the config file
        logging.debug("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
        logging.debug('Config loaded, program started')        
        return config
    except Exception as e:
        logging.debug('Failed to load config: ' + str(e))

'''
def load_keycombinations(config):
    try:
        # Do not include external files in the bundle security risk!!
        # Load the keycombinations file
        if getattr(sys, 'frozen', False):
            # Running in a bundle
            keycombinations_file_path = os.path.join(sys._MEIPASS, 'keycombinations_dict.json') # type: ignore
            logging.debug('Load keycombinations via sys._MEIPASS: ' + str(keycombinations_file_path))
        else:
            # Running in normal Python environment
            keycombinations_file_path = os.path.join(os.path.dirname(sys.argv[0]), 'keycombinations_dict.json')
            logging.debug('Load keycombinations via sys.argv[0]: ' + str(keycombinations_file_path))
        
        # redundant?
        #keycombinations_file_path = os.path.join(sys._MEIPASS, 'keycombinations_dict.json') # type: ignore
        #logging.debug('Load keycombinations via sys._MEIPASS: ' + str(keycombinations_file_path))
        # Load the keycombinations file
        with open(keycombinations_file_path) as keycombinations_file:
            keycombinations = json.load(keycombinations_file)
            logging.debug('Load keycombinations successfully!')

        return keycombinations
    except Exception as e:
        logging.debug('Failed to load keycombinations: ' + str(e))
'''


def local_keycombinations():
    try:
        return keycombinations
    except Exception as e:
        logging.debug('Failed to load local keycombinations: ' + str(e))

'''
def load_keycodes():
    
    try:   
        # Do not include external files in the bundle security risk!!
        # Load the keycodes file
        if getattr(sys, 'frozen', False):
            # Running in a bundle
            keycodes_file_path = os.path.join(sys._MEIPASS, 'keycodes.json') # type: ignore
            logging.debug('Load keycodes via sys._MEIPASS: ' + str(keycodes_file_path))
        else:
            # Running in normal Python environment
            keycodes_file_path = os.path.join(os.path.dirname(sys.argv[0]), 'keycodes.json')
            logging.debug('Load keycodes via sys.argv[0]: ' + str(keycodes_file_path))
        
        # redundant?
        # keycodes_file_path = os.path.join(sys._MEIPASS, 'keycodes.json') # type: ignore
        # logging.debug('Load keycodes via sys._MEIPASS: ' + str(keycodes_file_path))

        # Load the keycodes file
        with open(keycodes_file_path) as keycodes_file:
            keycodes = json.load(keycodes_file)
            logging.debug('Load keycodes successfully!')
        return keycodes
    except Exception as e:
        logging.debug('Failed to load keycodes: ' + str(e))
'''   
 
def local_keycodes():
    try:
        return keycodes
    except Exception as e:
        logging.debug('Failed to load local_keycodes: ' + str(e))

def open_serial_port(config):
    try:
        ser = serial.Serial(
            port=config['COM_PORT'],
            baudrate=config['BAUD_RATE'],
            bytesize=config['BYTESIZE'],
            parity=config['PARITY'],
            stopbits=config['STOPBITS'],
            xonxoff=config['XONXOFF'],
            rtscts=config['RTSCTS'],
            dsrdtr=config['DSRDTR']
        )
        logging.debug("Serial port " + str(config['COM_PORT']) + " opened successfully")
        logging.debug('-------------------------------------------------------------')
        return ser
    except serial.SerialException as e:
        logging.error('Failed to open serial port: ' + str(e))

def wait_for_window(config):
    try:
        if config['CHECK_FOR_WINDOW_ON_STARTUP']:
            logging.debug('Waiting for ' + str(config["WINDOW_NAME"]) + '...')
            window_handle = win32gui.FindWindow(None, config['WINDOW_NAME'])
            while window_handle == 0:
                time.sleep(0.3)
                window_handle = win32gui.FindWindow(None, config['WINDOW_NAME'])
            logging.debug('Found window handle: ' + str(window_handle))
            # return window_handle
    except Exception as e:
        logging.debug('Failed to get window handle: ' + str(e))

def get_window_handles(config):
    def enum_window_callback(hwnd, handles):
        if win32gui.GetWindowText(hwnd) == config['WINDOW_NAME']:
            handles.append(hwnd)   
    try:
        window_handles = []
        win32gui.EnumWindows(enum_window_callback, window_handles)
        
        log_once = True
        while not window_handles:
            if log_once:
                logging.debug('Dropped Window: ' + str(config['WINDOW_NAME']))
            log_once = False
            time.sleep(0.3)
            win32gui.EnumWindows(enum_window_callback, window_handles)
        
        return window_handles
    except Exception as e:
        logging.debug('Failed to get window handles: ' + str(e))

def get_z_order():
    hwnd_list = []
    def enum_windows_proc(hwnd, lParam):
        hwnd_list.append(hwnd)
        return True
    win32gui.EnumWindows(enum_windows_proc, 0)
    return hwnd_list

def find_lowest_target_window(config):
    window_handles = get_window_handles(config)
    z_list = get_z_order()
    lowest_hwnd = None
    for hwnd in z_list:
        if hwnd in window_handles:
            lowest_hwnd = hwnd
    return lowest_hwnd

def check_visible_z_order_above_target(config):
    lowest_hwnd = find_lowest_target_window(config)
    count = 0
    if lowest_hwnd is not None:
        hwnd_above = win32gui.GetWindow(lowest_hwnd, win32con.GW_HWNDPREV)
        while hwnd_above:
                if win32gui.IsWindowVisible(hwnd_above):
                    #'''
                    try:
                        window_text = win32gui.GetWindowText(hwnd_above)
                        if config['LOG_LEVEL'].upper() == 'DEBUG':
                            _, process_id = win32process.GetWindowThreadProcessId(hwnd_above)
                            try:
                                process = psutil.Process(process_id)
                                process_name = process.name()
                            except psutil.NoSuchProcess:
                                process_name = "Unknown"
                            logging.debug(f'Visible window above target:\t {process_name}\t\t{hwnd_above}\t{window_text}')
                    except Exception as e:
                        logging.debug('Failed to inspect Windows above Traget: ' + str(e))
                    #'''
                    if not (window_text == "") or config['SET_TO_FOREGROUND_STRICT']:
                        count += 1
                hwnd_above = win32gui.GetWindow(hwnd_above, win32con.GW_HWNDPREV)
        return count
    else:
        return count

def setFocusToTarget(config):
    # Function Gets called for each command reccived if the Foreground Window is not the Target Window
    # Function also gets called on regular intervals defined in the config file
    # Function checks if there are more visible windows in the Z-stack above the lowest Target Window then there are Target Windows - 1
    # (in addition to the lowest Target Window)
    # Function then Sets all the Target Windows to the foreground (brittle might break with Windows updates)
    window_handles = get_window_handles(config)
    if window_handles is None:
        logging.debug('No window handles found.')
        return None
    else:
        logging.debug('Window handles: ' + str(window_handles))
    
    # Get Z Satcking order and check if the first N number of windows match the list in window_handles (order is not important)
    # if not set focus to the windows missing in the first N number of windows in the Z stacking order but that are in window_handles
    # N is defined by Monitor_count in the config file

    
    # monitor_count = config['MONITOR_COUNT']
    # logging.debug('Monitor count: ' + str(monitor_count))
    
    visible_z_order_above_target = check_visible_z_order_above_target(config)
    logging.debug('Number of Target Window Handles: ' + str(len(window_handles)))
    logging.debug('Number of Windows above lowest Target: ' + str(visible_z_order_above_target))

    # Check if the first N windows in the Z order match the window_handles
    #z_order_top_n = z_order[:monitor_count]
    #logging.debug('Z order: ' + str(z_order))
    #missing_handles = [hwnd for hwnd in window_handles if hwnd not in z_order_top_n]
    logging.debug('Checking if any target Window is covert: ' + str((visible_z_order_above_target) > (len(window_handles) - 1)))
    if ((visible_z_order_above_target) > (len(window_handles) - 1)):
        #logging.debug('More Windows over lowest Target Window then other Target Windows: ' + str((visible_z_order_above_target) > (len(window_handles) - 1)))
        for hwnd in window_handles:
            try:
                attempts = 0
                max_attempts = 5
                while attempts < max_attempts:
                    try:
                        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                        win32gui.SetForegroundWindow(hwnd)
                        #win32gui.SetFocus(hwnd) # requires attachment to the thread
                        logging.debug('Set focus to window handle: ' + str(hwnd))
                        break
                    except Exception as e:
                        attempts += 1
                        logging.debug(f'Failed to set focus to window handle (attempt {attempts}): ' + str(e))
                        if attempts < max_attempts:
                            time.sleep(2)
            except Exception as e:
                logging.debug('Failed to set focus to window handle after multiple attempts: ' + str(e))
        time.sleep(0.5)
    return window_handles
    
def turn_off_caps_lock():
    if win32api.GetKeyState(win32con.VK_CAPITAL) & 1:
        time.sleep(0.05)
        win32api.keybd_event(win32con.VK_CAPITAL, 0, 0, 0)
        time.sleep(0.05)
        win32api.keybd_event(win32con.VK_CAPITAL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.05)

def send_keystrokes(command, keycodes, keycombinations, config):
    try:
        default_wait_time = config['WAIT_AFTER_KEY']
        keystroke_combination = keycombinations[command]
        for keystroke in keystroke_combination:
            if 'combination' in keystroke:          
                try:
                    key_combination = [keycodes[virtual_keycode] for virtual_keycode in keystroke['combination']]
                    time.sleep(0.05)
                    for virtual_keycode in key_combination:
                        win32api.keybd_event(virtual_keycode, 0, 0, 0)
                        time.sleep(0.05)
                        for keyResolve, value in keycodes.items():
                            if value == virtual_keycode:
                                logging.debug("win32api.keybd_event(" + str(keyResolve) + ", 0, 0, 0)")
                    for virtual_keycode in reversed(key_combination):
                        win32api.keybd_event(virtual_keycode, 0, win32con.KEYEVENTF_KEYUP, 0)
                        time.sleep(0.05)
                        for keyResolve, value in keycodes.items():
                            if value == virtual_keycode:
                                logging.debug("win32api.keybd_event(" + str(keyResolve) + ", 0, win32con.KEYEVENTF_KEYUP, 0)")
                except Exception as e:
                    logging.debug('Failed to send keystroke $COMBINATION: ' + str(e))
            elif 'press' in keystroke:
                try:
                    # Send a single key press
                    virtual_keycode = keycodes[keystroke['press']]
                    time.sleep(0.05)
                    win32api.keybd_event(virtual_keycode, 0, 0, 0)
                    time.sleep(0.05)
                    for keyResolve, value in keycodes.items():
                            if value == virtual_keycode:
                                logging.debug("win32api.keybd_event(" + str(keyResolve) + ", 0, 0, 0)")
                    win32api.keybd_event(virtual_keycode, 0, win32con.KEYEVENTF_KEYUP, 0)
                    time.sleep(0.05)
                    for keyResolve, value in keycodes.items():
                            if value == virtual_keycode:
                                logging.debug("win32api.keybd_event(" + str(keyResolve) + ", 0, win32con.KEYEVENTF_KEYUP, 0)")
                    
                except Exception as e:
                    logging.debug('Failed to send keystroke $PRESSED: ' + str(e))

            elif 'keep_pressed' in keystroke:
                # Hold down a key
                logging.debug("keep_pressed"+ str(keystroke))
                # When a Window is moved to the foreground while the key is being held down, the VMS is not focused again!!!! Maybe change this.
                try:
                    if keystroke['keep_pressed'] != "Release":
                        virtual_keycode = keycodes[keystroke['keep_pressed']]
                        time.sleep(0.05)
                        win32api.keybd_event(virtual_keycode, 0, 0, 0)
                        time.sleep(0.05)
                    else:
                        for data_set in keycombinations:
                            keystroke_combination = keycombinations[data_set]
                            for keystroke in keystroke_combination:
                                if 'keep_pressed' in keystroke and keystroke['keep_pressed'] != "Release":
                                    virtual_keycode = keycodes[keystroke['keep_pressed']]
                                    win32api.keybd_event(virtual_keycode, 0, win32con.KEYEVENTF_KEYUP, 0)
                                    time.sleep(0.05)
                except Exception as e:
                    logging.debug('Failed to handle keep_pressed keystroke: ' + str(e))
            if 'wait' in keystroke:
                wait_time = keystroke['wait'] if keystroke['wait'] != 'WAIT_AFTER_KEY' else default_wait_time
                time.sleep(wait_time)
    except Exception as e:
        logging.debug('Failed to send keystrokes: ' + str(e))

# Endlosschleife
# ADD INPUT Sanitizing!!!!
def listen_serial(ser, keycodes, keycombinations, config):
    try:
        terminator = b'a'
        # newline = b'\n'
        # command_buffer = ""
        
        last_check_time = 0  # Initialize last_check_time
        SET_TO_FOREGROUND_PAUSE = config['SET_TO_FOREGROUND_PAUSE']
        # window_handles = get_window_handles(config)
        
        while True:
            #window_handle = checkWindowHandle(config['WINDOW_NAME'])
            if ser.in_waiting > 0:
                command = ser.read_until(terminator).decode().strip()  # Decode the received bytes to a string
                logging.debug('Received command: {} <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<'.format(command))  # Add debug logging for received command
                if command:
                        try:
                            if command in keycombinations:
                                
                                try:
                                    if config['SET_TO_FOREGROUND_ENABLE']:
                                        # Check if currently focused window is part of the window_handles
                                        # If not run setFocusToTarget(config)
                                        window_handles = get_window_handles(config)
                                        focused_window = win32gui.GetForegroundWindow()
                                        last_check_time = time.time()
                                        logging.debug('Currently focused window handle: ' + str(focused_window))
                                        if not window_handles or focused_window not in window_handles:
                                            logging.debug('Focused window is not part of the target application windows, setting focus...')
                                            last_check_time = time.time()
                                            window_handles = setFocusToTarget(config)
                                except Exception as e:
                                    logging.debug('Failed to set focus to target: ' + str(e))
                                # Check if CAPS Lock is off!
                                turn_off_caps_lock()
                                send_keystrokes(command, keycodes, keycombinations, config)
                            else:
                                logging.warning('Command not found in keycombinations: {}'.format(command))
                        except Exception as e:
                            logging.debug('Failed to process command: ' + str(e))
            else:
                time.sleep(0.1)  # Add a small delay to reduce CPU usage

                try:
                    if config['SET_TO_FOREGROUND_ENABLE']:
                        current_time = time.time()
                        if current_time - last_check_time > SET_TO_FOREGROUND_PAUSE:
                            logging.debug('Automated Window Foreground Check:')
                            window_handles = setFocusToTarget(config)
                            last_check_time = current_time
                except Exception as e:
                    logging.debug('Failed to set focus to target on regular interval: ' + str(e))
    except Exception as e:
        logging.debug('Failed to listen to serial: ' + str(e))

time.sleep(0.05)

def main():
    try:
        enable_logging()
        config = load_config()
        #keycodes = load_keycodes()
        keycodes = local_keycodes()
        #keycombinations = load_keycombinations(config)
        keycombinations = local_keycombinations()
        time.sleep(5)
        wait_for_window(config)
        ser = open_serial_port(config)
        listen_serial(ser, keycodes, keycombinations, config)
    except Exception as e:
        logging.debug('Failed to execute main: ' + str(e))

if __name__ == "__main__":
    main()

