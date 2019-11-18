using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace G1ANT.Addon.LibreOffice
{
    public static class WriterManager
    {
    	private static List<WriterWrapper> launchedWriters = new List<WriterWrapper>();
        private static WriterWrapper currentWriter = null;

        public static WriterWrapper CurrentWriter
        {
            get
            {
                if (currentWriter == null)
                {
                    throw new ApplicationException("Writer instance must be opened first using writer.open command");
                }
                return currentWriter;
            }
            private set
            {
                currentWriter = value;
            }
        }

        public static void SwitchWriter(int id)
        {
           WriterWrapper instanceToSwitchTo = launchedWriters.Where(x => x.Id == id).FirstOrDefault();
           currentWriter = instanceToSwitchTo ?? throw new ArgumentException($"No Writer instance found with id: {id}");
        }

        private static int GetNextId()
        {
            return launchedWriters.Count() > 0 ? launchedWriters.Max(x => x.Id) + 1 : 0;
        }

        public static WriterWrapper CreateInstance()
        {
            int assignedId = GetNextId();
            WriterWrapper wrapper = new WriterWrapper(assignedId);
            launchedWriters.Add(wrapper);
            currentWriter = wrapper;
            return wrapper;
        }

        public static void RemoveInstance(int? id = null)
        {
            if (id == null)
            {
                id = currentWriter.Id;
            }
            var toRemove = launchedWriters.Where(x => x.Id == id).FirstOrDefault();
            if (toRemove != null)
            {
                launchedWriters.Remove(toRemove);
                toRemove.Close();
            }
            else
            {
                throw new ArgumentException($"Unable to close Writer instance with specified id argument: '{id}'");
            }
        }
    }
}
