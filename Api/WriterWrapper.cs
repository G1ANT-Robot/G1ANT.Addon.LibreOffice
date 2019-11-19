using System;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.text;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.util;

namespace G1ANT.Addon.LibreOffice
{
    public class WriterWrapper
    {
    	public int Id { get; set; }
        public WriterWrapper(int id)
        {
            this.Id = id;
        }

        public XTextDocument mxDocument;
        public XTextDocument MxDocument
        {
            get { return mxDocument ?? throw new ApplicationException("Writer instance must be opened first using calc.open command"); }
            private set { mxDocument = value; }
        }

        private XTextDocument InitDocument(bool hidden, string path)
        {
            var aLoader = CreateLoader();

            if (hidden)
            {
                var loadProperties = CreatePropertiesForHiddenReadonlyLoad();
                return IsNewDocument(path) ? CreateNewDocument(aLoader, loadProperties) : LoadDocument(path, aLoader, loadProperties);
            }
            else
            {
                return IsNewDocument(path) ? CreateNewDocument(aLoader) : LoadDocument(path, aLoader);
            }
        }

        private static bool IsNewDocument(string path)
        {
            return string.IsNullOrEmpty(path);
        }

        private static PropertyValue[] CreatePropertiesForHiddenReadonlyLoad()
        {
            return new PropertyValue[2] {
                new PropertyValue()
                {
                    Name = "Hidden",
                    Value = new uno.Any(true)
                },
                new PropertyValue()
                {
                    Name = "ReadOnly",
                    Value = new uno.Any(false)
                }
            };
        }

        private static XTextDocument LoadDocument(string path, XComponentLoader aLoader, PropertyValue[] loadProperties = null)
        {
            path = string.Format("file:///{0}", path);
            return (XTextDocument)aLoader.loadComponentFromURL(path, "_blank", 0, loadProperties ?? new PropertyValue[0]);
        }

        private static XTextDocument CreateNewDocument(XComponentLoader aLoader, PropertyValue[] loadProperties = null)
        {
            return (XTextDocument)aLoader.loadComponentFromURL("private:factory/swriter", "_blank", 0, loadProperties ?? new PropertyValue[0]);
        }

        private XComponentLoader CreateLoader()
        {
            var mxContext = uno.util.Bootstrap.bootstrap();
            var mxMSFactory = (XMultiServiceFactory)mxContext.getServiceManager();
            return (XComponentLoader)mxMSFactory.createInstance("com.sun.star.frame.Desktop");
        }

        private void SaveDocument(string path)
        {
            path = path.Replace("\\", "/"); // Convert forward slashes to backslashes, converting it to the correct format storeToURL expects. 
            path = string.Concat("file:///", path);
            var xStorable = (XStorable)MxDocument; // Typecast the currently open document to XStorable type.
            xStorable.storeToURL(path, new PropertyValue[1]); //Creating an empty PropertyValue array saves the document in the default .ods format.
        }

    	public void Close()
        {
            var xModel = (XModel)MxDocument;
            var xCloseable = (XCloseable)xModel;
            xCloseable.close(true);
        }

        public int Open(bool hidden, string path)
        {
            MxDocument = InitDocument(hidden, path);
            return this.Id;
        }

        public void Save(string path)
        {
            SaveDocument(path);
        }

        public void InsertText(string text , bool append)
        {
            XText xText = MxDocument.getText();
            if (append)
            {
                var xEnd = xText.getEnd();
                xEnd.setString(text);
            }
            else
            {
                xText.setString(text);
            }
        }

        public string GetText()
        {
            var xText = MxDocument.getText();
            return xText.getString();
        }
        
        public void ReplaceWith(string word, string replaceWith)
        {
            var xText = MxDocument.getText();
            var text = xText.getString();
            xText.setString(text.Replace(word, replaceWith));
        }
    }
}
