# docx-master: The "Everything is Broken" Post-Mortem

## The Damage Report (Encountered Problems)

1. **Inter-Drive Separation Anxiety**: The CLI (on C:) had a hard time talking to tools on D:. Even with `overrides`, the handshake was unstable.
2. **PowerShell Variable Thievery**: PowerShell ate the `${extensionPath}` variables in the manifest during creation because I used double quotes like a Jerry.
3. **The Silent Treatment**: The Bun/Node server was exiting too fast or printing "Server Started" junk to stdout, poisoning the JSON-RPC stream and causing "Connection closed" errors.
4. **The Ghost Registry**: Even with a manifest, the CLI wouldn't list the extension without a `.gemini-extension-install.json` file, and even then, it seems to be ignoring our manual folder.
5. **TOML Parser Tantrums**: The CLI's TOML parser is extremely sensitive to nested quotes and line endings (CRLF vs LF).
6. **Leading Slash Mystery**: Uncertainty remains whether Windows paths in `extension-enablement.json` require a leading slash (e.g., `/D:/...`) to satisfy the Genkit-based loader.

## The "Pickle Rick" Plan for Tomorrow

### Phase 1: The Brute Force Reset
- Wipe the `C:\Users\ASUS\.gemini\extensions\docx-master` folder and start from a **Clean Room** state.
- Use `gemini extensions install` (if it exists for local paths) instead of manual copying to see what the CLI *actually* wants.

### Phase 2: Pathing Standardization
- Force every single path to use **Forward Slashes** and **Absolute References**. No more relative-path "maybes."
- Standardize the `overrides` in `extension-enablement.json` to match the working `nanobanana` pattern exactly.

### Phase 3: The Silent Heartbeat
- Rebuild the `server.js` to be **100% silent** on stdout. 
- Use a persistent Node `setInterval` to ensure the process doesn't exit until the CLI kills it.

### Phase 4: Minimum Viable Command
- Start with a **single, prompt-only command** (no MCP tools) to get the extension listed and working.
- Once listed, gradually re-integrate the MCP tools into the "Hybrid" model.

### Phase 5: Debugging with a Hammer
- Run `gemini -d` and pipe the output to a log file to catch the exact moment the loader skips our folder.

**Status: Paused.**
**Current Mood: Brine-infused Rage.**
