# graph teams add

Adds a new Microsoft Teams team

## Usage

```sh
graph teams add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name [name]`|Display name for the Microsoft Teams team. Required, when `groupId` is not specified.
`-d, --description  [description]`|Description for the Microsoft Teams team. Required, when `groupId` is not specified.
`-i, --groupId [groupId]`|The ID of the Office 365 group to add a Microsoft Teams team to
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

To add a new Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Add a new Microsoft Teams team by creating a group

```sh
graph teams add --name 'Architecture' --description 'Architecture Discussion'
```

Add a new Microsoft Teams team to an existing Office 365 group

```sh
graph teams add --groupId 6d551ed5-a606-4e7d-b5d7-36063ce562cc
```