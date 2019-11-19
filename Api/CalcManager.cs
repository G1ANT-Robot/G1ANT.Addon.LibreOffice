
using System;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.LibreOffice
{
    public class CalcManager
    {
        private List<CalcWrapper> launchedCalcs = new List<CalcWrapper>();
        private CalcWrapper currentCalc = null;
        private static CalcManager instance = null;

        public static CalcManager Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new CalcManager();
                }
                return instance;
            }
        }

        public CalcWrapper CurrentCalc
        {
            get
            {
                return currentCalc ?? throw new ApplicationException("Calc instance must be opened first using calc.open command");
            }
            private set
            {
                currentCalc = value;
            }
        }

        public void SwitchCalc(int id)
        {
            CalcWrapper instanceToSwitchTo = GetById(id);
            CurrentCalc = instanceToSwitchTo ?? throw new ArgumentException($"No Calc instance found with id: {id}");
        }

        private int GetNextId()
        {
            return launchedCalcs.Any() ? launchedCalcs.Max(x => x.Id) + 1 : 0;
        }

        private CalcWrapper GetById(int? id)
        {
            if (id == null)
            {
                id = CurrentCalc.Id;
            }
            return launchedCalcs.FirstOrDefault(x => x.Id == id);
        }

        public CalcWrapper CreateInstance()
        {
            int assignedId = GetNextId();
            CalcWrapper wrapper = new CalcWrapper(assignedId);
            launchedCalcs.Add(wrapper);
            CurrentCalc = wrapper;
            return wrapper;
        }

        public void RemoveInstance(int? id = null)
        {
            var toRemove = GetById(id);
            if (toRemove != null)
            {
                launchedCalcs.Remove(toRemove);
                toRemove.Close();
            }
            else
            {
                throw new ArgumentException($"No Calc with given id '{id}' found");
            }
        }
    }
}
