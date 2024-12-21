import os
import json
import requests
import zipfile
import winreg
from rich.progress import Progress
from rich.console import Console
from rich.text import Text

USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
REGISTRY_PATH = 'SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall'
API_URL = 'https://api.tukui.org/v1/addon/elvui'

class WowInstallation:
	def __init__(self, name: str, path: str):
		self.name: str = self.update_name(name)
		self.path: str = path
		self.elvui_update_required: bool = False
		self.addons_path: str = f'{path}\\Interface\\AddOns'
		self.elvui_addon_path: str = f'{self.addons_path}\\ElvUI'

	def update_name(self, name: str) -> str:
		if name == 'World of Warcraft':
			return 'Retail'
		else:
			return name.removeprefix('World of Warcraft').strip()

	def elvui_get_toc(self) -> str:
		return self.elvui_addon_path + '\\ElvUI_Mainline.toc'

	def elvui_installed(self) -> bool:
		return os.path.exists(self.elvui_addon_path)

	def elvui_get_version(self) -> str | None:
		with open(self.elvui_get_toc(), 'r+') as f:
			for line in f.readlines():
				if line.startswith('## Version'):
					return line.split(': v')[1].strip()

		return None

installations : list[WowInstallation] = []
console = Console()

def cprint(s: str, end='\n') -> None:
	text = Text(s)
	text.highlight_words(['ElvUI'], style='blue')
	text.highlight_words(['✓'], style='green')
	text.highlight_words(['✕'], style='red')
	text.highlight_words(['→'], style='cyan')

	console.print(text, end=end)

def find_wow_installation_paths() -> None:
	with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, REGISTRY_PATH) as key:
		for i in range(winreg.QueryInfoKey(key)[0]):
			try:
				subkey = winreg.EnumKey(key, i)
				if 'World of Warcraft' in subkey:
					with winreg.OpenKey(key, subkey) as k:
						name = winreg.QueryValueEx(k, "DisplayName")[0]
						path = winreg.QueryValueEx(k, "DisplayIcon")[0]
						installations.append(WowInstallation(name, os.path.dirname(path)))
			except Exception:
				continue

def elvui_get_json() -> any:
	try:
		req = requests.get(API_URL)
		return edata.loads(req.content)
	except Exception:
		return None

def elvui_download(url: str, output_path: str) -> None:
	try:
		with requests.get(url, headers={'User-Agent': USER_AGENT}, stream=True, timeout=10, allow_redirects=True) as response:
			response.raise_for_status()
			size = int(response.headers.get('Content-Length', 0))

			with open(output_path, 'wb') as file:
				with Progress() as progress:
					task = progress.add_task(f'Downloading ElvUI {edata["version"]}', total=size)
					for chunk in response.iter_content(chunk_size=1024):
						if not chunk:
							break

						file.write(chunk)
						progress.update(task, advance=len(chunk))
	except requests.RequestException as e:
		cprint(f'✕ Error during download: {e}')
		exit()

def elvui_unzip(wow: WowInstallation, zip_path: str) -> None:
	with zipfile.ZipFile(zip_path, 'r') as zip:
		with Progress() as progress:
			task = progress.add_task(f'[{wow.name}] Extracting ElvUI', total=len(zip.namelist()))
			for file in zip.namelist():
				zip.extract(file, wow.addons_path)
				progress.update(task, advance=1)

# Locate all installed WoW versions from Windows registry.
find_wow_installation_paths()

# None of the WoW versions are currently installed.
if len(installations) == 0:
	cprint(f'✕ Could not locate WoW installation path(s).')
	exit()

# None of WoW versions have ElvUI installed.
if sum(wow.elvui_installed() for wow in installations) == 0:
	cprint('✕ None of WoW installation has ElvUI AddOn installed.')
	exit()

edata = elvui_get_json()

# Failed to fetch latest ElvUI version.
if edata is None:
	cprint('✕ Failed to fetch latest ElvUI version. Visit: https://tukui.org/elvui for manual download.')
	exit()

# Check for updates for each WoW version.
for wow in installations:
	if wow.elvui_installed():
		if edata['version'] != wow.elvui_get_version():
			wow.elvui_update_required = True
			cprint(f'→ [{wow.name}] ElvUI has a new update {edata["last_update"]}. Latest: {edata["version"]} (Installed: {wow.elvui_get_version()})')
		else:
			cprint(f'✓ [{wow.name}] ElvUI is up-to-date.')

# None of WoW versions need ElvUI update.
if not any(wow.elvui_update_required for wow in installations):
	exit()

# Ask user for updating all of the ElvUI installations.
update_prompt = input('Update all WoW versions to use the latest ElvUI? Y/N: ').lower()
if update_prompt != 'y':
	exit()

# Zip will be downloaded to current working directory.
zip_path = os.getcwd() + f'\\ElvUI-{edata["version"]}.zip'

elvui_download(edata['url'], zip_path)

if not os.path.exists(zip_path):
	exit()

for wow in installations:
	if wow.elvui_update_required:
		elvui_unzip(wow, zip_path)

# Finally remove the zip.
os.remove(zip_path)
cprint(f'✓ ElvUI updated!')
