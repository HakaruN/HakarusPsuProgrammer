﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HakarusKoradProgrammer
{
    public class TestSequenceElement
    {
        public float _voltage;
        public float _current;
        public int _time;
        public float _power;
        public float _resistance;
        public long _ElapsedMs;

        //This empty constructor is here purley for the Xml serializer, it needs a unperamiterised constructor.
        public TestSequenceElement()
        {

        }

        public TestSequenceElement(string voltage, string current, string time)
        {
            _voltage = float.Parse(voltage);
            _current = float.Parse(current);
            _time = int.Parse(time);
            _power = 0;
            _resistance = 0;

        }
        public TestSequenceElement(string voltage, string current, string power, string resistance, long ElapsedTime)
        {
            _voltage = float.Parse(voltage);
            _current = float.Parse(current);
            _power = float.Parse(power);
            _resistance = float.Parse(resistance);
            _ElapsedMs = ElapsedTime;
        }
        public float GetVoltage()
        {
            return _voltage;
        }
        public float GetCurrent()
        {
            return _current;
        }
        public int GetTime()
        {
            return _time;
        }
        public float GetPower()
        {
            return _power;
        }
        public float GetResistance()
        {
            return _resistance;
        }
        public long GetElapsedTime()
        {
            return _ElapsedMs;
        }
    }
}
