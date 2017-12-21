using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Education
{
    class Neums
    {
        private String neum;
        private String value;


        public Neums(String neum, String value)
        {
            this.Neum = neum;
            this.Value = value;
        }

        public string Neum { get => neum; set => neum = value; }
        public string Value { get => value; set => this.value = value; }
    }
}
