﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CivilTechReportGenerator.Interfaces {
    abstract class IDocument {

        public abstract void create();
        public abstract int count();
    }
}