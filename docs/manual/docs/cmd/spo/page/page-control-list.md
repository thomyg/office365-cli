# spo page control list

Lists controls on the specific modern page

## Usage

```sh
spo page control list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Name of the page to list controls of
`-u, --webUrl <webUrl>`|URL of the site where the page to retrieve is located
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To list controls on a modern page, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

If the specified name doesn't refer to an existing modern page, you will get a `File doesn't exists` error.

## Examples

List controls on the modern page with name _home.aspx_

```sh
spo page control list --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
```