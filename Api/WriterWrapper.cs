using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.bridge;
using unoidl.com.sun.star.frame;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice
{
    public class WriterWrapper
    {

    	public int Id { get; set; }

    	private unoidl.com.sun.star.uno.XComponentContext m_xContext;
        private unoidl.com.sun.star.lang.XMultiServiceFactory mxMSFactory;
        public unoidl.com.sun.star.text.XTextDocument mxDocument;

        public WriterWrapper(int id)
        {
            this.Id = id;
        }
        


        private XMultiServiceFactory Connect()
        {
            m_xContext = uno.util.Bootstrap.bootstrap();
            return (XMultiServiceFactory)m_xContext.getServiceManager();
        }

        private unoidl.com.sun.star.text.XTextDocument InitDocument(bool Hidden, String Path)
        {
            mxMSFactory = Connect();
            XComponentLoader aLoader = (XComponentLoader)mxMSFactory.createInstance("com.sun.star.frame.Desktop");
            XComponent xComponent;
            unoidl.com.sun.star.beans.PropertyValue[] loadProperties = new unoidl.com.sun.star.beans.PropertyValue[2];
            loadProperties[0] = new unoidl.com.sun.star.beans.PropertyValue();
            loadProperties[1] = new unoidl.com.sun.star.beans.PropertyValue();
            loadProperties[1].Name = "ReadOnly";
            loadProperties[1].Value = new uno.Any(false);

            if (Hidden)
            {
                loadProperties[0].Name = "Hidden";
                loadProperties[0].Value = new uno.Any(true);

                //Check to see if the path is provided by the user, if it's empty, create a new document. 
                if (String.IsNullOrEmpty(Path))
                {
                    xComponent = aLoader.loadComponentFromURL("private:factory/swriter", "_blank", 0, loadProperties);
                }
                else
                {
                    try
                    {
                        Path = String.Format("file:///{0}", Path);
                        xComponent = aLoader.loadComponentFromURL(Path, "_blank", 0, loadProperties);
                    }
                    catch (unoidl.com.sun.star.uno.Exception ex)
                    {
                        throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
                    }
                }
            }
            else
            {
                loadProperties[0].Name = "Hidden";
                loadProperties[0].Value = new uno.Any(false);
                if (String.IsNullOrEmpty(Path))
                {
                    xComponent = aLoader.loadComponentFromURL("private:factory/swriter", "_blank", 0, new unoidl.com.sun.star.beans.PropertyValue[0]);
                }
                else
                {
                    try
                    {
                        Path = String.Format("file:///{0}", Path);
                        xComponent = aLoader.loadComponentFromURL(Path, "_blank", 0, new unoidl.com.sun.star.beans.PropertyValue[0]);
                    }
                    catch (unoidl.com.sun.star.uno.Exception ex)
                    {
                        throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
                    }
                }
            }

            return (unoidl.com.sun.star.text.XTextDocument)xComponent;
        }

        private void SaveDocument(String Path)
        {
            try
            {
                Path = Path.Replace("\\", "/"); // Convert forward slashes to backslashes, converting it to the correct format storeToURL expects. 
                Path = String.Concat("file:///", Path);
                XStorable xStorable = (XStorable)mxDocument; // Typecast the currently open document to XStorable type.
                xStorable.storeToURL(Path, new unoidl.com.sun.star.beans.PropertyValue[1]); //Creating an empty PropertyValue array saves the document in the default .ods format.
            }
            catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }
        }


    	public void Close()
        {
            try
            {
                XModel xModel = (XModel)mxDocument;
                unoidl.com.sun.star.util.XCloseable xCloseable = (unoidl.com.sun.star.util.XCloseable)xModel;
                xCloseable.close(true);
            }
            catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }
        }

        public int Open(bool Hidden, String Path)
        {
            mxDocument = InitDocument(Hidden, Path);
            return this.Id;
        }

        public void Save(String Path)
        {
            SaveDocument(Path);
        }

        public void InsertText(String text , bool append)
        {
            try
            {
                unoidl.com.sun.star.text.XText xText = mxDocument.getText();
                if (append)
                {
                    unoidl.com.sun.star.text.XTextRange xEnd = xText.getEnd();
                    xEnd.setString(text);
                }
                else
                {
                    xText.setString(text);
                }
            }
            catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }

        }

        public String GetText()
        {
            unoidl.com.sun.star.text.XText xText = mxDocument.getText();
            return xText.getString();
        }
        
        public void ReplaceWith(String word, String replaceWith)
        {
            unoidl.com.sun.star.text.XText xText = mxDocument.getText();
            var text = xText.getString();
            xText.setString(text.Replace(word, replaceWith));
        }
    }
}
