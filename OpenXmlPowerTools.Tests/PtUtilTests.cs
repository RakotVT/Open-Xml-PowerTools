﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class PtUtilTests
    {
        [Theory(Skip = "This is failing on AppVeyor")]
        [InlineData("PU/PU001-Test001.mht")]
        public void PU001(string name)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceMht = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var src = File.ReadAllText(sourceMht.FullName);
            var p = MhtParser.Parse(src);
            Assert.True(p.ContentType != null);
            Assert.True(p.MimeVersion != null);
            Assert.True(p.Parts.Length != 0);
            Assert.DoesNotContain(p.Parts, part => part.ContentType == null || part.ContentLocation == null);
        }

    }
}

#endif
