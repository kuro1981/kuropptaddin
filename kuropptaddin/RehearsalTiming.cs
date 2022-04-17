using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace kuropptaddin
{
    public class RehearsalTiming
    {

        private PowerPoint.Slide _slide;
        private String[] _timings;

        public PowerPoint.Slide Slide
        {
            set { _slide = value; }
        }

        public String[] Timings {
            get { return _timings; }
            set { _timings = value; }
        }

        public RehearsalTiming(PowerPoint.Slide slide)
        {
            _slide = slide;
            _timings = getTimingTag();
        }

        private string[] getTimingTag()
        {
            bool found = false;
            string timing = "";


            for (int i = 1; i <= _slide.Tags.Count; i++)
            {
                if (_slide.Tags.Name(i) == "TIMING")
                {
                    found = true;
                    timing = _slide.Tags.Value(i);
                    Console.WriteLine(timing);
                    break;
                }
            }

            return timing.Split('|');

        }
    }
}
