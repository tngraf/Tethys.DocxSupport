# ---------------------------------------------
# Clean project
# SPDX-FileCopyrightText: (c) 2022-2023 T. Graf
# SPDX-License-Identifier: Apache-2.0
# ---------------------------------------------

dotnet clean
Remove-Item "Tethys.DocxSupport\bin" -Recurse
Remove-Item "Tethys.DocxSupport\obj" -Recurse