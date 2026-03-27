# Conversation History Downloader

Python script that lists ElevenLabs ConvAI conversations for your agents, downloads full conversation details from the API, and writes one Excel workbook per agent. Optional JSON state avoids repeat GET requests for conversation IDs you have already downloaded.

## Requirements

- Python 3.10 or newer (3.13 recommended)
- An ElevenLabs API key with access to ConvAI conversation APIs

## Setup

1. Clone this repository.

2. Create a virtual environment and install dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

On Windows, use `.venv\Scripts\activate` instead of `source .venv/bin/activate`.

## Configuration

Set your API key in the environment before running the script:

```bash
export ELEVENLABS_API_KEY="your_key_here"
```

The key is read from `ELEVENLABS_API_KEY` or `XI_API_KEY`.

Optional: set `ELEVENLABS_BASE_URL` if your account uses a regional API host listed in the ElevenLabs documentation.

## Usage

Run from the project directory with the virtual environment activated.

Interactive mode (you will be prompted for one or more agent IDs, comma separated):

```bash
python export_conversations.py
```

Non-interactive mode:

```bash
python export_conversations.py --agent-id YOUR_AGENT_ID
```

Multiple agents (one Excel file per agent):

```bash
python export_conversations.py --agent-id AGENT_A --agent-id AGENT_B -o ./output
```

## Command line options

- `--agent-id` Agent ID. Repeat the flag for multiple agents. If omitted, the script prompts for IDs.
- `-o` or `--output-dir` Directory for Excel files. Default is the current directory.
- `--base-url` API base URL. Default is `https://api.elevenlabs.io` unless `ELEVENLABS_BASE_URL` is set.
- `--page-size` Page size for the list endpoint (1 to 100). Default 30.
- `--max-pages` Maximum number of list API pages to fetch per agent. Each page returns up to `--page-size` conversations. Use `0` for no limit (default). Stops early if the API reports more pages but the cap is reached, so exports may be partial.
- `--pause` Seconds to sleep between list pagination and GET calls. Default 0.2.
- `--state-file` Path to the JSON file that stores fetched conversation IDs. Default is `.elevenlabs_captured_ids.json` inside `--output-dir`.
- `--no-state` Do not read or write state. Every run calls GET for every conversation returned by the list API.
- `--force-refresh` Ignore saved state and GET every conversation again. State is still updated after new fetches.

Use `python export_conversations.py --help` for full help text.

## Output

Each run produces one workbook per agent, named like:

`conversations_<agent_id>_<UTC_timestamp>.xlsx`

Nested JSON fields in API responses become column names with dot notation. Array values are stored as JSON strings in a single cell.

## State file and API usage

The default state file records which conversation IDs have already been retrieved with GET. On later runs, those IDs are skipped for GET calls to reduce API usage. The list endpoint is still called each run to discover IDs.

Conversation IDs that are skipped are not included in that run workbook. Older runs keep their own Excel files.

If the state file is corrupt JSON, the script prints a warning and starts with an empty cache.

## Security

Do not commit API keys or `.env` files that contain secrets. The `.gitignore` is configured to ignore common local artifacts and exports.

## License

Add a license file if you distribute this project.
