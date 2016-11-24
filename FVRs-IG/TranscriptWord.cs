using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FVRs_IG
{
    class TranscriptWord
    {
        public string Name { get; set; }
        public int Frequency { get; set; }
        public List<Occurrence> PageAndLine { get; set; }

        public TranscriptWord()
        {
            PageAndLine = new List<Occurrence>();
        }

    }
}
