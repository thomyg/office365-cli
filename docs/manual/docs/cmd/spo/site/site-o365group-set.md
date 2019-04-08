# spo site o365group set

Connects site collection to an Office 365 Group

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

## Usage

```sh
spo site o365group set [options]
```

## Alias

```sh
spo site groupify
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --siteUrl <siteUrl>`|URL of the site collection being connected to new Office 365 Group
`-a, --alias <alias>`|The email alias for the new Office 365 Group that will be created
`-n, --displayName <displayName>`|The name of the new Office 365 Group that will be created
`-d, --description [description]`|The group’s description
`-c, --classification [classification]`|The classification value, if classifications are set for the organization. If no value is provided, the default classification will be set, if one is configured
`--isPublic`|Determines the Office 365 Group’s privacy setting. If set, the group will be public, otherwise it will be private
`--keepOldHomepage`|For sites that already have a modern page set as homepage, set this option, to keep it as the homepage
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To connect site collection to an Office 365 Group, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

When connecting site collection to an Office 365 Group, SharePoint will create a new group using the specified information. If a group with the same name already exists, you will get a `The group alias already exists.` error.

## Examples

Connect site collection to an Office 365 Group

```sh
spo site o365group set --siteUrl https://contoso.sharepoin.com/sites/team-a --alias team-a --displayName 'Team A'
```

Connect site collection to an Office 365 Group and make the group public

```sh
spo site o365group set --siteUrl https://contoso.sharepoin.com/sites/team-a --alias team-a --displayName 'Team A' --isPublic
```

Connect site collection to an Office 365 Group and set the group classification

```sh
spo site o365group set --siteUrl https://contoso.sharepoin.com/sites/team-a --alias team-a --displayName 'Team A' --classification HBI
```

Connect site collection to an Office 365 Group and keep the old home page

```sh
spo site o365group set --siteUrl https://contoso.sharepoin.com/sites/team-a --alias team-a --displayName 'Team A' --keepOldHomepage
```

## More information

- Overview of the "Log in to new Office 365 group" feature: [https://docs.microsoft.com/en-us/sharepoint/dev/features/groupify/groupify-overview](https://docs.microsoft.com/en-us/sharepoint/dev/features/groupify/groupify-overview)
