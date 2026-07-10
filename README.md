# URL to IP Resolver

A Python utility that extracts hostnames from a list of URLs, resolves them to IP addresses, and exports the results to an Excel spreadsheet with duplicate IP addresses highlighted.

## Features

* Reads URLs from a text file.
* Extracts the hostname from each URL.
* Removes the protocol (`http://` or `https://`) and any URL path.
* Resolves each hostname to one or more IP addresses.
* Exports the results to an Excel (`.xlsx`) file.
* Highlights duplicate IP addresses.
* Uses a different color for each duplicate IP group.
* Automatically adjusts column widths.

## Example

### Input (`urls.txt`)

```text
https://example.com/login
http://sub.example.com/admin
https://google.com/search?q=test
github.com
```

### Output

| Original URL                     | Hostname        | IP Address    |
| -------------------------------- | --------------- | ------------- |
| https://example.com/login        | example.com     | 93.184.216.34 |
| http://sub.example.com/admin     | sub.example.com | 93.184.216.34 |
| https://google.com/search?q=test | google.com      | 142.250.x.x   |
| github.com                       | github.com      | 140.82.x.x    |

Duplicate IP addresses are automatically highlighted using different colors.

## Requirements

* Python 3.8+
* openpyxl

## Installation

Clone the repository:

```bash
https://github.com/anmolbagul/URLs_to_IP_Resolver.git
cd URLs_to_IP_Resolver
```

(Optional) Create a virtual environment:

```bash
python3 -m venv venv
```

Activate the virtual environment.

**Linux/macOS**

```bash
source venv/bin/activate
```

**Windows**

```powershell
venv\Scripts\activate
```

Install the dependencies:

```bash
pip install openpyxl
```

## Usage

```bash
python3 resolve_urls.py -i urls.txt -o nslookup_results.xlsx
```

### Arguments

| Argument         | Description                                                    |
| ---------------- | -------------------------------------------------------------- |
| `-i`, `--input`  | Input text file containing URLs (required).                    |
| `-o`, `--output` | Output Excel file (optional). Default: `nslookup_results.xlsx` |

### Example

```bash
python3 resolve_urls.py -i urls.txt -o results.xlsx
```

## Notes

* Multiple IP addresses returned for a hostname are recorded as separate rows.
* Hostnames that cannot be resolved are marked as **Could not resolve**.
* Duplicate IP addresses are highlighted automatically.
* DNS resolution uses the system resolver (`socket.gethostbyname_ex()`).

## License

MIT License
