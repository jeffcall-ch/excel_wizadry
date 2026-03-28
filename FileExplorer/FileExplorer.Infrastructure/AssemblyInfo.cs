// Copyright (c) FileExplorer contributors. All rights reserved.

// Trimming descriptor for AOT-sensitive types.
// Include this file in AOT publish profiles to preserve COM interop types.

using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("FileExplorer.Tests")]
