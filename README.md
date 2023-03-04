<!-- 
SPDX-FileCopyrightText: (c) 2022-2023 T. Graf
SPDX-License-Identifier: Apache-2.0
-->

# Tethys.DocxSupport

![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)
[![Build status](https://ci.appveyor.com/api/projects/status/3k0fy06set3or784?svg=true)](https://ci.appveyor.com/project/tngraf/tethys-docxsupport)
[![Nuget](https://img.shields.io/badge/nuget-1.0.0-brightgreen.svg)](https://www.nuget.org/packages/Tethys.DocxSupport/1.0.0)
[![REUSE status](https://api.reuse.software/badge/git.fsfe.org/reuse/api)](https://api.reuse.software/info/git.fsfe.org/reuse/api)

This library simplifies working with the [Open XML SDK for Office](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk?redirectedfrom=MSDN).
The **DocX** format has an enormous number of features, but if you just want to create a simple document
there is a steep learning curve. Tethys.DocxSupport simplies a number of operations.

## Get Package

You can get Tethys.DocxSupport by grabbing the latest NuGet packages from [here](https://www.nuget.org/packages/Tethys.DocxSupport/1.0.0).

## Build

### Requisites

* Visual Studio 2019

### Build Solution

Just use the basic `dotnet` command:

```shell
dotnet build
```

Run the demo application:

```shell
dotnet run --project .\Tethys.DocxSupport.Demo\Tethys.DocxSupport.Demo.csproj
```

## License

Tethys.DocxSupport is licensed under the Apache License, Version 2.0.
