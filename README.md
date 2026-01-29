<!-- toc:insertAfterHeading=PSVSCodeExtensionPolicyRoH-->
<!-- toc:insertAfterHeadingOffset=4-->

# PSVSCodeExtensionPolicyRoH

_This module imports a function to manage VSCode extensions via PowerShell on windows._


## Table of Contents

1. [Introduction](#introduction)
1. [Getting started](#getting-started)
    1. [Prerequisites](#prerequisites)
    1. [Installation](#installation)
1. [How to use](#how-to-use)
    1. [How to Import](#how-to-import)
    1. [Using the function](#using-the-function)
1. [LICENSE](#license)

## Introduction

Because of emerging threads in supply chains and dev environments, i decided to create a PowerShell module which can be used to manage VSCode extensions. This module can be used via Intune or any oder MDM to manage VSCode extensions globally on a windows device or in usercontext. The recommended way is systemcontext. 

A documentation can be found [here][VSCodeExtensionManagement].

## Getting started

### Prerequisites

- Powershell installed, works with Windows Powershell (preinstalled on Windows) and Powershell Core
- Operatingsystem: Windows
- IDE like VS Code, if you want to contribute or change the code

### Installation

The module is not published to PSGallery so you can only download it from github:

1. Using Git:

```PS
# Powershell
# Pull necessary files.
git clone "https://github.com/IT-Administrators/PSVSCodeExtensionPolicyRoH.git"
# Change location to project directory.
cd .\PSVSCodeExtensionPolicyRoH\
```

2. Using Powershell

```PS
# Download zip archive to current directory using powershell.
Invoke-WebRequest -Uri "https://github.com/IT-Administrators/PSVSCodeExtensionPolicyRoH/archive/refs/heads/main.zip" -OutFile "PSVSCodeExtensionPolicyRoH.zip"
# Than expand archive.
Expand-Archive -Path ".\PSVSCodeExtensionPolicyRoH.zip"
```

## How to use

### How to Import

You can import the module in two ways:

1. Import from current directory 
```PS
# Import from current directory
Import-Module -Path ".\PSVSCodeExtensionPolicyRoH.psm1" -Force -Verbose
```
2. Copy it to your module directory to get it imported on every session start. This depends also on your executionpolicy.

To get your module directorys use the following command:

```PS
$env:PSModulePath
```

### Using the function

After the module is imported, the function ```Set-VSCodeExtensionPolicy``` will be available.

The following shows how to configure the allowed extensions for all users on the current device:

```PS
# Set allowed publishers of VSCode extensions.
Set-VSCodeExtensionPolicy -AllowedPublishers "microsoft"-SystemContext -Verbose

# Registry values set
extensions.allowed          : {"microsoft":true}
extensions.autoUpdate       : 0
extensions.autoCheckUpdates : 0
extensions.gallery.enabled  : 1
```

## LICENSE

[MIT][License]


[VSCodeExtensionManagement]: https://code.visualstudio.com/docs/enterprise/extensions
[License]: ./LICENSE