# VMS Interface

## Warning

This program enables the remote sending of keystrokes via the Windows API through COM Port or MQTT, either generally or directly to a specific application. 
This poses a significant security risk if not implemented correctly. 
The application is still under development. 
Proceed with CAUTION!

**DO NOT USE IN A PRODUCTION ENVIRONMENT!!!**


## Installation

### Setup Python Enviromnt

1.  Make Sure Python 3.12 is installed on the Maschine (does not need to be added to PATH)
2.  Create Virtual Enviroment in this Directory using:
	`<Full-Path-to-Python.exe> -m venv .venv`
    Like this:  
    `C:\Users\Username\py312\python.exe -m venv .venv`
    (The name .venv is importatnt because it is assumend in: vscode-settings.json: `"python.defaultInterpreterPath":".venv/Scripts/python.exe"`)
3.  activate the venv using:
	  `activate` in pwsh Terminal  
    Or if that does not work use::  
    `cd .venv/Scripts && ./activate && cd ../..`
4.  install all required python packages (Internet Connection)  
    `pip install -r requirements.txt`
    confirm installed packages with  
    `pip list`
5.  put the config.yaml in: `C:\ProgramData\VMSInterface`

#### requirements.txt contens:
```txt
psutil==7.0.0
pyserial==3.5
pywin32==306
PyYAML==6.0.2
pyinstaller==6.8.0
paho-mqtt==2.1.0
```


### publish .exe:

`cd .venv/Scripts && ./activate && cd ../.. ` 
`pyinstaller main.spec`

**.exe will be in /dist**

## License

Weingarten (Baden), Germany FEB. 2025

**Lead Developer:** [github.com/emsysol](https://github.com/emsysol)  
Parts of this code were generated using **GitHub Copilot AI** under GitHub 
user account: **[github.com/batchdeploy](https://github.com/batchdeploy)**

----------------------------------------------------------------------------

**Disclaimer for keycombinations_dict.py:**

The original `keycombinations_dict.py` is not part of the GPL v3 
Licensed Project since it is HMI specific and possibly contains commands 
that are under NDA. It needs to be newly created by the user. 
Refer to the `keycombinations_dict_manual.md` for reference and how to.

----------------------------------------------------------------------------

This program is released under the GNU General Public License v3.0.

You are free to use, modify, and distribute this software under the terms of  
the GNU GPL v3, either version 3 or any later version.

This program is distributed in the hope that it will be useful,  
but WITHOUT ANY WARRANTY; without even the implied warranty of  
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the  
GNU General Public License for more details.

You should have received a copy of the GNU General Public License  
along with this program. If not, see [https://www.gnu.org/licenses/](https://www.gnu.org/licenses/).

## Author

**Lead Developer:** [github.com/emsysol](https://github.com/emsysol)  
Parts of this code were generated using **GitHub Copilot AI** under the GitHub user account: **[github.com/batchdeploy](https://github.com/batchdeploy)**

## Disclaimer:

THIS SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, AND NON-INFRINGEMENT.  
IN NO EVENT SHALL THE AUTHORS, COPYRIGHT HOLDERS, OR CONTRIBUTORS BE LIABLE FOR ANY CLAIM, DAMAGES, OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT, OR OTHERWISE, ARISING FROM, OUT OF, OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

**USE THIS SOFTWARE AT YOUR OWN RISK.** No guarantees are made regarding functionality, security, or performance.

Parts of this code were generated using **GitHub Copilot AI** under the GitHub user account: **[github.com/batchdeploy](https://github.com/batchdeploy)**.

By using this software, you acknowledge that portions were generated with the assistance of **GitHub Copilot AI** and accept the associated terms and conditions.

## Acknowledgments

### Built using:

- [PySerial](https://github.com/pyserial/pyserial) – A Python library for serial port communication.
- [PyWin32](https://github.com/mhammond/pywin32) – Python extensions for Windows.
- [PyYAML](https://github.com/yaml/pyyaml) – A YAML parser and emitter for Python.

**See THIRD_PARTY_NOTICES for more Information.**