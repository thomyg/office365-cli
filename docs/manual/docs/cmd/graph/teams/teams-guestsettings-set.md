# graph teams guestsettings set

Updates guest settings of a Microsoft Teams team

## Usage

```sh
graph teams guestsettings set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the Teams team for which to update settings
`--allowCreateUpdateChannels [allowCreateUpdateChannels]`|Set to `true` to allow guests to create and update channels and to `false` to disallow it
`--allowDeleteChannels [allowDeleteChannels]`|Set to `true` to allow guests to create and update channels and to `false` to disallow it
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To update guest settings of the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Allow guests to create and edit channels

```sh
graph teams guestsettings set --teamId '00000000-0000-0000-0000-000000000000' --allowCreateUpdateChannels true
```

Disallow guests to delete channels

```sh
graph teams guestsettings set --teamId '00000000-0000-0000-0000-000000000000' --allowDeleteChannels false
```