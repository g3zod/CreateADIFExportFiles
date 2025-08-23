#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace AdifExportFilesCreator
{
    internal class XhtmlFileLoader
    {
        private static readonly byte[] Utf8Bom = [0xEF, 0xBB, 0xBF];

        private static readonly Encoding Windows1252Encoding = CodePagesEncodingProvider.Instance.GetEncoding(1252)
            ?? throw new InvalidOperationException("Windows-1252 encoding is not available.");

        /**
         * <summary>
         *   This method loads an XHTML file into an XmlDocument object.<br />
         *   <br />
         *   There are limitations suited to dealing with the specific XHTML files the program is designed for.<br />
         *   - The file is Windows-1252, UTF-8, or ISO-88591-1 encoded.<br />
         *   - The file is well-formed XHTML set up for HTML backwards compatibility including a text/html mime type.<br />
         *   - The only character entity supported is nbsp (non-breaking space).
         * </summary>
         *
         * <remarks>
         *   For the ADIF Specification, versions up to and including 3.1.5 are Windows-1252 and versions 3.1.6 onwards are UTF-8.
         * </remarks>
         * 
         * <param name="xhtmlFilePath">The location of the XHTML file.</param>
         * 
         * <returns>An XmlDocument object containing the contents of the XHTML file.</returns>
         */
        internal static XmlDocument Load(string xhtmlFilePath)
        {
            string tempDocPath = Path.GetTempFileName();
            XmlDocument xmlDocIn = new();

            try
            {
                xmlDocIn.PreserveWhitespace = true;
                Encoding xhtmlFileEncoding;

                // Determine whether the file is Windows-1252 is or UTF-8.

                using (FileStream s = new(xhtmlFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    byte[] buffer = new byte[Utf8Bom.Length];

                    if (s.Read(buffer, 0, buffer.Length) != Utf8Bom.Length)
                    {
                        throw new FileLoadException($"Failed to read first {Utf8Bom.Length} bytes from {xhtmlFilePath} to determine encoding", xhtmlFilePath);
                    }
                    xhtmlFileEncoding = buffer.SequenceEqual(Utf8Bom)
                        ? Encoding.UTF8
                        : Windows1252Encoding;
                }

                using (StreamReader htmlDocInStream = new(xhtmlFilePath, xhtmlFileEncoding))
                {
                    // Re-write the document as a true XML file and add an internal DTD to replace &nbsp; character
                    // entities with a space.  (To be pedantic, it should replace them with &#160; but there is
                    // no point because this program replaces all white-space sequences with a single space anyway.)

                    using StreamWriter tempDocStream = new(tempDocPath, false, xhtmlFileEncoding);

                    {
                        // Skip <html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">

                        string? line;

                        do
                        {
                            line = htmlDocInStream.ReadLine();
                            if (line == null)
                            {
                                throw new FileLoadException($"Failed to find <html> tag in {xhtmlFilePath}", xhtmlFilePath);
                            }
                        } while (!line.Contains("<html"));
                    }

                    tempDocStream.WriteLine($"<?xml version=\"1.0\" encoding=\"{xhtmlFileEncoding.WebName}\" ?>");
                    tempDocStream.WriteLine("<!DOCTYPE html [ <!ENTITY nbsp \" \"> ]>");
                    tempDocStream.WriteLine("<html>");

                    while (!htmlDocInStream.EndOfStream)
                    {
                        tempDocStream.WriteLine(htmlDocInStream.ReadLine());
                    }
                }

                // Passing xmlDocIn.Load() a file containing Windows-1252 encoding will fail because it cannot deal with it.
                // To overcome this, it is necessary to set up a StreamReader with a Windows-1252 encoding object and
                // use that with xmlDocIn.Load()
                //
                // While that's not a problem for files that are UTF-8 encoded, the same code using a StreamReader will work fine.

                using StreamReader sw = new(tempDocPath, xhtmlFileEncoding);
                xmlDocIn.Load(sw);

                string contentAttribute = string.Empty;

                XmlNodeList? metaNodeList = xmlDocIn?.SelectNodes("html/head/meta");

                if (metaNodeList != null)
                {
                    foreach (XmlNode metaNode in metaNodeList)
                    {
                        if (metaNode is XmlElement metaElement)
                        {
                            // If the meta element has an http-equiv attribute, it must be "Content-Type" or "content-type".
                            // If it has a content attribute, it must be "text/html; charset=windows-1252" or "text/html; charset=utf-8".
                            //
                            // E.g. <meta http-equiv="content-type" content="text/html; charset=utf-8" />

                            if (metaElement.HasAttribute("http-equiv") &&
                                metaElement.GetAttribute("http-equiv").Equals("Content-Type", StringComparison.OrdinalIgnoreCase) &&
                                metaElement.HasAttribute("content"))
                            {
                                contentAttribute = metaElement.GetAttribute("content");
                                break;
                            }
                        }
                    }
                }

                DecodeContentType(contentAttribute, out string _, out string contentTypeCharset);

                // Ensure the encoding declared in the XML file matches whether or not the file has a UTF-8 byte order mark.
                // This could happen if the XHTML file declares itself as UTF-8 but has been saved without the byte order mark.

                if (contentTypeCharset == "windows-1252")
                {
                    if (xhtmlFileEncoding != Windows1252Encoding)
                    {
                        throw new FileLoadException($"The encoding in the XHTML file is \"{contentAttribute}\" but the file is UTF-8 encoded", xhtmlFilePath);
                    }
                }
                else if (contentTypeCharset == "utf-8")
                {
                    if (xhtmlFileEncoding != Encoding.UTF8)
                    {
                        throw new FileLoadException($"The encoding in the XHTML file is \"{contentAttribute}\" but the file is not UTF-8 encoded");
                    }
                }
                else if (contentTypeCharset == "iso-8859-1")
                {
                    // ISO-8859-1 is a subset of UTF-8, so no check is necessary; if the file is UTF-8 encoded, it will be fine and
                    // otherwise it's impractical to discover whether the file itself is actually ISO-8859-1 encoded.
                }
                else
                {
                    throw new FileLoadException($"The encoding in the XHTML file {contentAttribute} is not Windows-1252 or UTF-8", xhtmlFilePath);
                }
            }
            finally
            {
                try
                {
                    File.Delete(tempDocPath);
                }
                catch { }
            }
            return xmlDocIn;
        }

        /**
         * <summary>
         *   Takes a content-type string and extracts the media type and character set.  This is used with the content-type value
         *   from a meta tag.  E.g.<br />
         *   <![CDATA[  <meta http-equiv="content-type" content="text/html; charset=utf-8" />]]>
         * </summary> 
         *
         * <param name="contentType">The content type string e.g. "text/html; charset=windows-1252".</param>
         * <param name="media">The media type in lower case, e.g. "text/html" if one is included, otherwise an empty string.</param>
         * <param name="charset">If a charset was included, its value in lower case, e.g. "utf-8", otherwise an empty string.</param>
         */
        private static void DecodeContentType(
            string contentType,
            out string media,
            out string charset)
        {
            // E.g. "text/html; charset=windows-1252"

            media = string.Empty;
            charset = string.Empty;
            string[] parts = contentType.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            if (parts.Length > 0)
            {
                media = parts[0].ToLowerInvariant();
            }
            foreach (string part in parts)
            {
                if (part.StartsWith("charset=", StringComparison.OrdinalIgnoreCase))
                {
                    charset = part["charset=".Length..].Trim().ToLowerInvariant();
                }
            }
        }
    }
}
