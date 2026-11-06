@echo off
chcp 65001 >nul
REM Claude Code external tools: graphify, ponytail, caveman, ralph-loop, commit-commands, claude-md-management, claude-code-setup, frontend-design, headroom, notion, hwp-mcp
REM Needs: Node.js, Python 3.10+, Git, Claude Code CLI

where claude >nul 2>nul || (
  echo [!] claude not found. Installing Claude Code CLI...
  call npm install -g @anthropic-ai/claude-code
)

echo === graphify ===
pip install graphifyy
graphify install || python -m graphify install

echo === ponytail ===
call claude plugin marketplace add DietrichGebert/ponytail
call claude plugin install ponytail@ponytail

echo === caveman ===
call claude plugin marketplace add JuliusBrussee/caveman
call claude plugin install caveman@caveman

echo === ralph-loop ===
call claude plugin install ralph-loop@anthropic-plugin-directory || (
  call claude plugin marketplace add anthropics/claude-plugins-official
  call claude plugin install ralph-loop@claude-plugins-official
)

echo === Anthropic official plugins ===
for %%P in (commit-commands claude-md-management claude-code-setup frontend-design) do (
  call claude plugin install %%P@anthropic-plugin-directory || call claude plugin install %%P@claude-plugins-official
)

echo === headroom (directory only) ===
call claude plugin install headroom@anthropic-plugin-directory || echo [!] headroom not found - skip (use it in claude.ai app)

echo === notion ===
call claude plugin install notion@anthropic-plugin-directory || (
  call claude plugin marketplace add makenotion/claude-code-notion-plugin
  call claude plugin install notion
)
echo [i] notion: in Claude Code run /mcp and log in once

echo === Python libraries ===
pip install xlwings ezdxf pyproj

echo === hwp-mcp (Hangul COM) ===
if not exist "%~dp0hwp-mcp" git clone --depth 1 https://github.com/jkf87/hwp-mcp.git "%~dp0hwp-mcp"
pip install -r "%~dp0hwp-mcp\requirements.txt"
pip install --force-reinstall pywin32
call claude mcp add hwp -s user -- python "%~dp0hwp-mcp\hwp_mcp_stdio_server.py"
call claude mcp list

echo.
echo === installed plugins ===
call claude plugin list
pause
