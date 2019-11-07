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

        private unoidl.com.sun.star.sheet.XSpreadsheetDocument initDocument()
        {
            XComponentLoader aLoader = (XComponentLoader)mxMSFactory.createInstance("com.sun.star.frame.Desktop");
            XComponent xComponent = aLoader.loadComponentFromURL(
                "private:factory/scalc", "_blank", 0,
                new unoidl.com.sun.star.beans.PropertyValue[0]);
            return (unoidl.com.sun.star.sheet.XSpreadsheetDocument)xComponent;
        }

        private XMultiServiceFactory connect(String[] args)
        {
            m_xContext = uno.util.Bootstrap.bootstrap();
            return (XMultiServiceFactory)m_xContext.getServiceManager();
        }

        public int Open()
        {
            unoidl.com.sun.star.sheet.XSpreadsheetDocument mxDocument = initDocument();
            return this.Id;
        }
    }
}
