using System;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.LibreOffice
{
    public class WriterManager
    {
    	private List<WriterWrapper> launchedWriters = new List<WriterWrapper>();
        private WriterWrapper currentWriter = null;
        private static WriterManager instance = null;

        private WriterManager() { }

        public static WriterManager Instance
        {
            get
            {
                if(instance == null)
                {
                    instance = new WriterManager();
                }
                return instance;
            }
        }

        public WriterWrapper CurrentWriter
        {
            get
            {
                return currentWriter ?? throw new ApplicationException("Writer instance must be opened first using writer.open command");
            }
            private set
            {
                currentWriter = value;
            }
        }

        public void SwitchWriter(int id)
        {
           WriterWrapper instanceToSwitchTo = GetById(id);
           currentWriter = instanceToSwitchTo ?? throw new ArgumentException($"No Writer instance found with id: {id}");
        }

        private int GetNextId()
        {
            return launchedWriters.Any() ? launchedWriters.Max(x => x.Id) + 1 : 0;
        }

        private WriterWrapper GetById(int? id)
        {
            if (id == null)
            {
                id = currentWriter.Id;
            }
            return launchedWriters.FirstOrDefault(x => x.Id == id);
        }

        public WriterWrapper CreateInstance()
        {
            var assignedId = GetNextId();
            var wrapper = new WriterWrapper(assignedId);
            launchedWriters.Add(wrapper);
            currentWriter = wrapper;
            return wrapper;
        }

        public void RemoveInstance(int? id = null)
        {
            var toRemove = GetById(id);
            if (toRemove != null)
            {
                launchedWriters.Remove(toRemove);
                toRemove.Close();
            }
            else
            {
                throw new ArgumentException($"No Writer with given id '{id}' found");
            }
        }
    }
}
