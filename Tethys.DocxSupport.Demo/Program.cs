// ---------------------------------------------------------------------------
// <copyright file="BasicWordSupport.cs" company="Tethys">
//   Copyright (C) 2022-2023 T. Graf
// </copyright>
//
// Licensed under the Apache License, Version 2.0.
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
// either express or implied.
// SPDX-License-Identifier: Apache-2.0
// ---------------------------------------------------------------------------

namespace Tethys.DocxSupport.Demo
{
    using System;

    /// <summary>
    /// Main class of the application.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        public static void Main(string[] args)
        {
            Console.WriteLine("Creating document...");
            var doc = new MyDocument();
            doc.CreateFullDocument("MyDocument.docx");
            Console.WriteLine("Done.");
        }
    }
}
