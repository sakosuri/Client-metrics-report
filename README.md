# Catalyst Center Client Metrics Report

A Python command-line tool that retrieves per-client wireless metrics from Cisco Catalyst Center and exports a multi-sheet Excel report.

The report combines:

- Five-minute trend analytics for RSSI, SNR, rates, onboarding duration, and roaming duration
- Point-in-time client details for health score and connected access point information
- Assurance events with client details sampled at each event timestamp
- Bar charts summarizing onboarding, roaming, RSSI, and SNR distributions

## Requirements

- Python 3.9 or newer
- Access to a Cisco Catalyst Center appliance
- Catalyst Center credentials with permission to read client analytics, client details, and assurance events
- An Excel workbook containing client MAC addresses

Install the Python dependencies from the repository directory:

```bash
python3 -m pip install -r requirements.txt
```

## Configuration

Create a `.env` file in the same directory as `client_metrics_report.py`:

```dotenv
CATC_API_BASE=https://catalyst-center.example.com
CATC_AUTH_URL=https://catalyst-center.example.com/dna/system/api/v1/auth/token
CATC_USERNAME=your_username
CATC_PASSWORD=your_password
```

Configuration variables:

| Variable | Description |
| --- | --- |
| `CATC_API_BASE` | Catalyst Center base URL, without a trailing slash |
| `CATC_AUTH_URL` | Catalyst Center authentication endpoint |
| `CATC_USERNAME` | Catalyst Center username |
| `CATC_PASSWORD` | Catalyst Center password |

Do not commit `.env` files or generated reports containing operational data to source control. The script currently disables TLS certificate verification because Catalyst Center installations commonly use self-signed certificates. Use a trusted certificate and enable verification before deploying this tool in a security-sensitive environment.

## Input workbook

By default, the script reads `client_metrics_report_input.xlsx` from the current working directory. Client MAC addresses must be in the first column, one address per row. A header row is optional.

Example:

| MAC Address |
| --- |
| `aa:bb:cc:dd:ee:ff` |
| `11:22:33:44:55:66` |

The following first-column headers are ignored automatically:

- `MAC Address`
- `MAC`
- `Client MAC`
- `MACAddress`

## Usage

Run the script from this directory:

```bash
python3 client_metrics_report.py --time-range 24h --input clients.xlsx
```

Generate a seven-day report with a specific output filename:

```bash
python3 client_metrics_report.py \
  --time-range 7d \
  --input clients.xlsx \
  --output client_metrics_7d.xlsx
```

If `--time-range` is omitted, the script prompts for either the last 24 hours or the last 7 days:

```bash
python3 client_metrics_report.py --input clients.xlsx
```

### Command-line options

| Option | Default | Description |
| --- | --- | --- |
| `--time-range` | Prompted | `24h` or `7d` |
| `--input` | `client_metrics_report_input.xlsx` | Input `.xlsx` file with MAC addresses in column A |
| `--output` | `client_metrics_report_DD_MM_YY.xlsx` | Destination report filename |

The selected time window ends at the time the report starts. Timestamps in the generated workbook are displayed in IST (`UTC+05:30`).

## Output workbook

The generated workbook contains four sheets:

### Client Metrics Summary

One row per input client. Trend metrics are the average of all available five-minute maximum samples in the selected window.

Columns include:

- RSSI and SNR
- Onboarding and roaming duration in seconds
- Tx and Rx rate in kBps, and data rate in kbps
- Health score
- Connected device MAC and name

### 5-Min Interval Data

Raw five-minute trend samples after unit conversion. This sheet includes the timestamp, RSSI, SNR, durations, and rates for each client and interval.

### Client Events

Assurance events returned for each client during the selected window. For each event, the script requests client details at the event timestamp and records event-specific metrics where available.

### Charts

Four embedded bar charts:

- Client onboarding times
- Client roaming times
- Connectivity RSSI
- Connectivity SNR

The charts include threshold shading and a percentage summary based on the clients with available values.

## Catalyst Center APIs

The script calls these endpoints:

- `POST /dna/system/api/v1/auth/token` for authentication
- `POST /dna/data/api/v1/clients/{mac}/trendAnalytics` for five-minute client trend data
- `GET /dna/intent/api/v1/client-detail` for point-in-time client details
- `GET /dna/data/api/v1/assuranceEvents` for client assurance events

Trend analytics requests use a five-minute interval and request maximum aggregates for RSSI, SNR, Tx/Rx rate, data rate, onboarding duration, and roaming duration. Trend results are paginated using the response cursor.

## Unit conversions

The report applies these conversions:

- Durations: milliseconds to seconds
- Tx/Rx rates: bytes per second to kilobytes per second (kBps) by dividing by `1000`
- Data rate: bits per second to kilobits per second (kbps) by dividing by `1000`

Values that are missing or cannot be converted are left blank in the workbook.

## Troubleshooting

### Authentication fails

Check that all four environment variables are present, that the authentication URL is reachable, and that the credentials can access Catalyst Center APIs.

### No metrics are returned

Verify the MAC address format, confirm that the client existed during the selected window, and check that the Catalyst Center user has access to client trend analytics.

### The input file is not found

Pass an explicit path with `--input` or run the command from the directory containing `client_metrics_report_input.xlsx`.

### API warnings appear

The script prints warnings for unsuccessful details, events, or trend requests and continues processing the remaining clients. Review the HTTP status and response details printed in the terminal.

## Security notes

- Keep credentials in environment variables or a local `.env` file, never in source code.
- Add `.env` and generated `.xlsx` reports to `.gitignore when using this tool in a Git repository.
- Review TLS verification settings before using the script outside a controlled internal environment.
- Treat generated reports as potentially sensitive because they contain client identifiers, health data, and network device information.
