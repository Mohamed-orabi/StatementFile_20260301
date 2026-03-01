using System;
using System.IO;
using System.Xml;


/// Summary description for clsXmlReadWrite.
public class clsXmlReadWrite
{


  public void WriteData()
  {
    string m_strFileName = String.Empty;
    XmlTextWriter bankWriter = null;
    bankWriter = new XmlTextWriter(m_strFileName, null);

    try
    {
      bankWriter.Formatting = Formatting.Indented;
      bankWriter.Indentation = 6;
      bankWriter.Namespaces = false;

      bankWriter.WriteStartDocument();

      bankWriter.WriteStartElement(String.Empty, "Parms", String.Empty);

      bankWriter.WriteStartElement(String.Empty, "UserName", String.Empty);
      bankWriter.WriteString("mUserName");
      bankWriter.WriteEndElement();

      bankWriter.WriteStartElement(String.Empty, "Password", String.Empty);
      bankWriter.WriteString("mPassword");
      bankWriter.WriteEndElement();

      bankWriter.WriteStartElement(String.Empty, "Server", String.Empty);
      bankWriter.WriteString("mServer");
      bankWriter.WriteEndElement();

      bankWriter.WriteEndElement();
      bankWriter.Flush();
    }
    catch (Exception e)
    {
      Console.WriteLine("Exception: {0}", e.ToString());
    }
    finally
    {
      if (bankWriter != null)
      {
        bankWriter.Close();
      }
    }
  }



  private void LoadData()
  {
    // Create an isntance of XmlTextReader and call Read method to read the file
    string m_strFileName = "D:\\pC#\\__\\C#\\Network\\_Email\\VB_MailChecker\\parms.xml";
    XmlTextReader bankReader = null;
    bankReader = new XmlTextReader(m_strFileName);

    while (bankReader.Read())
    {
      if (bankReader.NodeType == XmlNodeType.Element)
      {
        if (bankReader.LocalName.Equals("UserName"))
        {
          //mUserName = bankReader.ReadString();
        }

        if (bankReader.LocalName.Equals("Password"))
        {
          //mPassword = bankReader.ReadString();
        }

        if (bankReader.LocalName.Equals("Server"))
        {
          //mServer = bankReader.ReadString();
        }
      }
    }
    bankReader.Close();

  }



  public static void readConfigration()
  {
    XmlTextReader reader = null;
    try
    {

      //Create the reader.
      reader = new XmlTextReader(@Environment.CurrentDirectory + "\\" + clsBasUserData.applicationName + ".xml");
      //reader.WhitespaceHandling = WhitespaceHandling.None;

      //Parse the XML document.  ReadString is used to 
      //read the text content of the elements.
      reader.Read();
      reader.ReadStartElement("Settings");
      clsDbCon.sPackage = reader.ReadElementString("Package");
      clsDbCon.sServer = reader.ReadElementString("ServerName");
      clsDbCon.sUserName = reader.ReadElementString("UserName");
      clsDbCon.sPassword = reader.ReadElementString("Password");
      //reader.ReadEndElement();

    }
    catch (NotSupportedException ex)  //(Exception ex) 
    {
      clsBasErrors.catchError(ex);
    }
    finally
    {
      //Close the reader.
      reader.Close();
    }


  }

  public static void writeConfigration()
  {
    XmlTextWriter writer = null;
    System.IO.StreamWriter xmlSW = null;
    try
    {

      //Create the reader.
      xmlSW = new System.IO.StreamWriter(@Environment.CurrentDirectory + "\\" + clsBasUserData.applicationName + ".xml", true, System.Text.Encoding.ASCII);
      writer = new XmlTextWriter(xmlSW);
      writer.Formatting = Formatting.Indented;
      //reader.WhitespaceHandling = WhitespaceHandling.None;

      writer.WriteStartElement("Stock");
      writer.WriteAttributeString("Symbol", "Symbol1");
      writer.WriteElementString("Price", "Price1");
      writer.WriteElementString("Change", "Change1"); //XmlConvert.ToString(
      writer.WriteElementString("Volume", "Volume1");
      writer.WriteEndElement();

    }
    catch (NotSupportedException ex)  //(Exception ex) 
    {
      clsBasErrors.catchError(ex);
    }
    finally
    {
      //Close the reader.
      writer.Close();
      xmlSW.Close();
    }

  }
  /*
  */


  public clsXmlReadWrite()
  {
    //
    // TODO: Add constructor logic here
    //
  }
}

