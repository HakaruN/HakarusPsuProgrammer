using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HakarusKoradProgrammer
{
    class TestSequence
    {
        private List<TestSequenceElement> _TestSequenceElement = new List<TestSequenceElement>();
        public TestSequence()
        {
        }

        public void SetTestSequence(List<TestSequenceElement> TestSequenceElements)
        {
            _TestSequenceElement = TestSequenceElements;
        }

    }
}
