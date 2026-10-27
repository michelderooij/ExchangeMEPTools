# Copilot Instructions

## Project

This repository is a PowerShell module for analysing and clearing named properties from mailboxes in on-premises Exchange.

## Working Guidelines

- Keep changes focused on the module's mailbox named-property workflows.
- Follow the PowerShell style and compatibility requirements established by the module's existing code; do not assume a PowerShell or Exchange version without evidence.
- Inspect project scripts and tests before choosing validation commands. Use `./build.ps1` and the standalone assertion script; do not introduce Pester.
- Avoid adding dependencies or changing public behavior unless the task requires it.