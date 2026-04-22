# Smath-Printers

<img width="256" height="256" alt="smath_printers_logo" src="https://github.com/mguihard77/Smath-Printers/blob/main/smath_printers_logo.png" />

Smath Printers is a Powershell based utility to easily deploy printers in your domain.

With all the issues and limitations encoutered with Microsoft GPO Printers Deployement, this little program allows you to configure your printers in a .csv file then to simply deploy them on your domain computers.

Compilated using <a href="https://github.com/MScholtes/PS2EXE" target="_blank">PS2EXE-GUI</a>


---

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

All behavior is driven by a single CSV file named `config.csv`, placed next to the script or executable.

The CSV defines **both printer removals and printer installations**.

---

## Removal rules (REMOVE)

Printers can be removed explicitly using the `REMOVE` section of the CSV.

Removal rules are applied **before any installation**, and behave exactly like the original script logic.

### Supported removal modes

| Mode | Description |
|------|-------------|
| `PREFIX` | Removes all printers whose name starts with the given value |
| `EXACT`  | Removes a printer matching the exact name |

Associated TCP/IP ports (`IP_*`) are also removed.

---

## Example CSV File

```csv
Section,Mode,Value,Name,IP,Port,Driver,InfPath,Type
REMOVE,PREFIX,Office-Printer_,,,,,,
REMOVE,EXACT,Old_Printer_Name,,,,,,

Name,IP,Port,Driver,InfPath,Type
Office_Printer_1,192.168.1.10,IP_192.168.1.10,Xerox Global Print Driver PCL6,"\\server\drivers\xerox\driver.inf",Xerox
Secure_Print,127.0.0.1,IP_127.0.0.1,Xerox Global Print Driver PCL6,,Print2Me
Floor2_Printer,192.168.1.30,IP_192.168.1.30,Microsoft PS Class Driver,,Sharp
```

---

## Printer types

Smath-Printers supports a limited set of **predefined printer types**.

These values are not free-form: they are directly mapped to explicit behaviors inside the script. This keeps the deployment logic simple and predictable.

### Supported types

| Type value | Behavior |
|-----------|----------|
| `Xerox`   | Standard Xerox printer installation |
| `Print2Me`| Logical printer using a local TCP/IP port (127.0.0.1) |
| `Sharp`   | Disables bidirectional support (`EnableBidi = false`) |

Other values are not supported and will be ignored.

---

## Usage

PowerShell script

```csv
.\smath_printers.ps1
```

Compiled executable

When compiled as an executable, config.csv must be located in the same directory as the executable.

---

## Permissions

Smath-Printers is designed to run without administrator privileges in standard deployment scenarios.

The script:

- does not restart the print spooler
- does not modify the Windows registry
- does not perform system-level cleanup

As long as the execution context has permission to manage local printers, the deployment will work correctly.
