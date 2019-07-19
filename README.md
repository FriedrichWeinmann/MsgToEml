# MsgToEml

## Description

A small project designed to convert MSG Email files to EML format.

## Note

This project utilizes EWS and Exchange to perform the conversion, so EWS is needed.

It also uses Outlook COM to load the Outlook-Proprietary MSG file into the mailbox

## Example

```powershell
Get-ChildItem *.msg | Convert-MsgToEml
```

This will convert all msg files in the current folder to eml (or die trying).