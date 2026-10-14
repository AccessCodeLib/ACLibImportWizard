# ACLib-Import-Wizard
Import code from [AccessCodeLib](https://github.com/AccessCodeLib/AccessCodeLib) into application by add-in

Details: https://wiki.access-codelib.net/ACLib-Import-Wizard

#### Using the Import-Wizard:
[Video: create new Access application](https://access-codelib.net/videos/ACLib-Import-Wizard/neue-anwendung-erstellen/)

[Video: Import from GitHub repository](https://access-codelib.net/videos/ACLib-Import-Wizard/ACLib-Import-Wizard_ImportFromGitHub.mp4)

#### Automation API
The wizard can be used without UI, e.g. from scripts or AI agents.
For AI agents (MCP clients like Claude Code) there is an MCP server for this API:
[ACLibImportWizard-MCP](https://github.com/AccessCodeLib/ACLibImportWizard-MCP)

`API` returns a JSON string:

```vba
Application.Run "<add-in path without .accda>.API", "<Command>", "<Files>", "<Options>"

' Example (installed add-in)
Application.Run Environ("APPDATA") & "\Microsoft\AddIns\ACLibImportWizard.API", _
                "Resolve", "form/filter/FilterControlManager.cls", "mode=missing"
```

| Command | Arguments | Result |
|---|---|---|
| `GetVersion` | - | add-in and API version |
| `GetConfig` | [Options] | configured repositories, resulting source |
| `ListInstalled` | - | ACLib modules in the current database (`<file>` from the codelib block) |
| `Resolve` | Files, [Options] | read only: files incl. dependencies, references, tests, planned action (`import`/`skip`/`overwrite`), missing files |
| `Import` | Files, [Options] | imports like the wizard form; imported/skipped files, missing dependencies, added/removed references, messages |

* Files: `"data/SqlTools.cls;text/String2.cls"`
* Options: `"mode=missing|all|selected;tests=0|1;examples=0|1;source=config|local|github;repo=owner/name@branch"`
  (defaults: `mode=missing`, `tests=0`, `examples=0`, `source=config` = local repository if configured, otherwise GitHub `AccessCodeLib/AccessCodeLib@master`)

Result: `{"ok":true,...}` or `{"ok":false,"error":{"number":...,"source":"...","description":"..."}}` (no message boxes).

```json
{"ok":true,"source":"local","repository":"C:\CodeLib\AccessCodeLib\\","mode":"missing","tests":false,"examples":false,
 "files":[{"name":"String2","file":"text/String2.cls","type":"class","action":"import"},
          {"name":"_AccessCodeLib_license","file":"_codelib/license.bas","type":"module","action":"skip"},
          {"name":"StringTools","file":"text/StringTools.bas","type":"module","action":"import"}],
 "missing":[],"references":{"added":[],"removed":[]},"executes":[],"messages":[]}
```

Notes:
* The add-in and the target database must be in trusted locations. Otherwise Access shows a
  security prompt (the call blocks) or disables code.
* `Import` stops before importing if a selected file is not available. Missing dependencies are
  skipped (as in the wizard form) and listed in `missing`.
* Compile the project after `Import` and check the result.
* With [msaccess-vcs-mcp](https://github.com/joyfullservice/msaccess-vcs-mcp):
  `vcs_call_vba(database_path="<target>.accdb", function_name="<add-in path without .accda>.API", args=["Resolve", "text/String2.cls"])`

#### Contributing
See [CONTRIBUTING.md](CONTRIBUTING.md) for branches, releases and development.
