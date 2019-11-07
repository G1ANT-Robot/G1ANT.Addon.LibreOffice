
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace G1ANT.Addon.LibreOffice
{
    public static class CalcManager
    {

        private static List<CalcWrapper> launchedCalcs = new List<CalcWrapper>();
        private static CalcWrapper currentCalc = null;

        public static CalcWrapper CurrentCalc
        {
            get
            {
                if (currentCalc == null)
                {
                    throw new ApplicationException("Excel instance must be opened first using excel.open command");
                }
                return currentCalc;
            }
            private set
            {
                currentCalc = value;
            }
        }

        public static void SwitchExcel(int id)
        {
            CalcWrapper instanceToSwitchTo = launchedCalcs.Where(x => x.Id == id).FirstOrDefault();
            if (instanceToSwitchTo == null)
            {
                throw new ArgumentException($"No Calc instance found with id: {id}");
            }
            CurrentCalc = instanceToSwitchTo;
        }

        private static int GetNextId()
        {
            return launchedCalcs.Count() > 0 ? launchedCalcs.Max(x => x.Id) + 1 : 0;
        }

        public static CalcWrapper CreateInstance()
        {
            int assignedId = GetNextId();
            CalcWrapper wrapper = new CalcWrapper(assignedId);
            launchedCalcs.Add(wrapper);
            CurrentCalc = wrapper;
            return wrapper;
        }

        public static void RemoveInstance(int? id = null)
        {
            if (id == null)
            {
                id = CurrentCalc.Id;
            }
            var toRemove = launchedCalcs.Where(x => x.Id == id).FirstOrDefault();
            if (toRemove != null)
            {
                launchedCalcs.Remove(toRemove);
                //toRemove.Close();
            }
            else
            {
                throw new ArgumentException($"Unable to close Calc instance with specified id argument: '{id}'");
            }
        }
    }
}
