# Smath-Printers

<img width="256" height="256" alt="smath_printers_logo" src="https://github.com/mguihard77/Smath-Printers/blob/main/smath_printers_logo.png" />

Smath Printers is a Powershell based utility to easily deploy printers in your domain.

With all the issues and limitations encoutered with Microsoft GPO Printers Deployement, this little program allows you to configure your printers in a .csv file then to simply deploy them on your domain computers.

Compilated using <a href="https://github.com/MScholtes/PS2EXE" target="_blank">PS2EXE-GUI</a>

## Why Smath-Printers?

Deploying printers through Group Policy Objects can quickly become unreliable:

- random deployment failures  
- duplicated printer queues  
- orphan TCP/IP ports  
- inconsistent behavior across computers  
- difficult troubleshooting  

Smath-Printers was designed as a **simple and deterministic alternative**, focused on reliability rather than complexity.

---

## How it works

The workflow intentionally follows the same logic every time:

1. Existing printers matching the deployment scope are removed  
2. Associated TCP/IP ports are also removed  
3. Printers defined in `config.csv` are recreated from scratch  
4. A graphical progress bar is displayed during execution  
5. A final summary shows what has been removed and installed  

No registry cleanup, no print spooler manipulation, no state detection.

---

## CSV configuration

Printers are fully defined using a single CSV file named `config.csv`, placed next to the script or executable.

### Example `config.csv`

```csv
Name,IP,Port,Driver,InfPath,Type
Printer_Floor1,192.168.1.10,IP_192.168.1.10,Generic Printer Driver,"\\server\drivers\generic.inf",Generic
Printer_Floor2,192.168.1.20,IP_192.168.1.20,Generic Printer Driver,,Generic
Printer_Color,192.168.1.30,IP_192.168.1.30,Microsoft PS Class Driver,,Sharp
``
