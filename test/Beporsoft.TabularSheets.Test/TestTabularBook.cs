﻿using Beporsoft.TabularSheets.Test.Helpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test
{
    [Category("Workbook")]
    public class TestTabularBook
    {
        private readonly bool _clearFolderOnEnd = false;
        private readonly int _amountRows = 1000;
        private readonly TestFilesHandler _filesHandler = new();
        private readonly Stopwatch _stopwatch = new Stopwatch();

        [SetUp]
        public void Setup()
        {
            _stopwatch.Reset();
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            if (_clearFolderOnEnd)
                _filesHandler.ClearFiles();

        }


        [Test, Ignore("Not implemented yet")]
        public void BookAddSheet_ShouldAppendSheetAtEnd()
        {

        }


        [Test, Ignore("Not implemented yet")]
        public void BookRemoveSheet_ShouldReturnFalse_IfNotIncluded()
        {

        }


        [Test, Ignore("Not implemented yet")]
        public void BookRemoveSheet_ShouldReturnTrueAndReallocateOthers()
        {

        }


        [Test, Ignore("Not implemented yet")]
        public void CreateMultiple_ShouldIterateSheetName_IfSameTableName()
        {

        }

        [Test, Ignore("Not implemented yet")]
        public void CreateMultiple_ShouldIncludeEach_AsExpected()
        {

        }

        [Test, Ignore("Not implemented yet")]
        public void CreateMultiple_ShouldShareStrings()
        {

        }

        [Test, Ignore("Not implemented yet")]
        public void CreateMultiple_ShouldShareStyles_IfEqual()
        {
        }

    }
}
