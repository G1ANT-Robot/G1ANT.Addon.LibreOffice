using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.bridge;
using unoidl.com.sun.star.frame;

namespace G1ANT.Addon.LibreOffice
{
    public class CalcWrapper
    {
        public int Id { get; set; }

        public CalcWrapper(int id)
        {
            this.Id = id;
        }

        private unoidl.com.sun.star.uno.XComponentContext m_xContext;
        private unoidl.com.sun.star.lang.XMultiServiceFactory mxMSFactory;
        public unoidl.com.sun.star.sheet.XSpreadsheetDocument mxDocument;

        private unoidl.com.sun.star.sheet.XSpreadsheetDocument InitDocument(bool Hidden, String Path)
        {
            mxMSFactory = Connect();
            XComponentLoader aLoader = (XComponentLoader)mxMSFactory.createInstance("com.sun.star.frame.Desktop");
            XComponent xComponent;

            if (Hidden)
            {
                unoidl.com.sun.star.beans.PropertyValue[] loadProperties = new unoidl.com.sun.star.beans.PropertyValue[1];
                loadProperties[0] = new unoidl.com.sun.star.beans.PropertyValue();
                loadProperties[0].Name = "Hidden";
                loadProperties[0].Value = new uno.Any(true);
                //Check to see if the path is provided by the user, if it's empty, create a new document. 
                if (String.IsNullOrEmpty(Path))
                {
                    xComponent = aLoader.loadComponentFromURL("private:factory/scalc", "_blank", 0, loadProperties);
                }
                else
                {
                    try
                    {
                        Path = String.Format("file:///{0}", Path);
                        xComponent = aLoader.loadComponentFromURL(Path, "_blank", 0, loadProperties);
                    }
                    catch
                    {
                        throw new unoidl.com.sun.star.io.IOException($"Failed to open file [{Path}]", this);
                    }
                }
            }
            else
            {
                if (String.IsNullOrEmpty(Path))
                {
                    xComponent = aLoader.loadComponentFromURL("private:factory/scalc", "_blank", 0, new unoidl.com.sun.star.beans.PropertyValue[0]);
                }
                else
                {
                    try
                    {
                        Path = String.Format("file:///{0}", Path);
                        xComponent = aLoader.loadComponentFromURL(Path, "_blank", 0, new unoidl.com.sun.star.beans.PropertyValue[0]);
                    }
                    catch
                    {
                        throw new unoidl.com.sun.star.io.IOException($"Failed to open file: [{Path}]", this);
                    }
                }
            }

            return (unoidl.com.sun.star.sheet.XSpreadsheetDocument)xComponent;
        }

        private XMultiServiceFactory Connect()
        {
            m_xContext = uno.util.Bootstrap.bootstrap();
            return (XMultiServiceFactory)m_xContext.getServiceManager();
        }

        public int Open(bool Hidden, String Path)
        {
            mxDocument = InitDocument(Hidden, Path);
            return this.Id;
        }
    }
}
