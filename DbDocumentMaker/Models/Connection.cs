﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbDocumentMaker.Models
{
    public class Connection
    {
        // Properties
        public string Name { get; set; }
        
        public string Str { get; set; }
        public string SystemName { get; set; } = string.Empty;
        public string SystemDescription { get; set; } = string.Empty;


        // Methods
        public override string ToString()
        {
            return Name;
        }

        public bool IsValid()
        {
            return string.IsNullOrWhiteSpace(this.Name) == false && string.IsNullOrWhiteSpace(this.Str) == false;
        }
    }
}
