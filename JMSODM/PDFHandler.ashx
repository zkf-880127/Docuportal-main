<%@ WebHandler Language="C#" Class="PDFHandler" %>

using System;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using iTextSharp.text;
using System.IO;
using Webapps.Utils;

public class PDFHandler : IHttpHandler, System.Web.SessionState.IRequiresSessionState
{

    public void ProcessRequest(HttpContext context)
    {
        int id, sourceId, ImageArchived;
        string fileName = string.Empty;
        string fileType = string.Empty;
        try
        {
            id = int.Parse(context.Request.QueryString["Id"]);
            sourceId = int.Parse(context.Request.QueryString["SId"]);
            ImageArchived = int.Parse(context.Request.QueryString["ImageArchived"]);//ImageArchived  zkf 2012107-9 update archive database 
        }
        catch {
            byte[]Bytes = GetPDFBytesFromTextIsNull();
            context.Response.Buffer = true;
            context.Response.Charset = "";
            context.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            context.Response.ContentType = "application/pdf";
            context.Response.BinaryWrite(Bytes);
            context.Response.Flush();
            context.Response.End();
            return;
        }


        try
        {
            if (!CustomRoles.RolesForPageLoad())
            {
                return;
            }
        }
        catch (Exception ex)
        {
            return;
        }
        finally
        { }

        //ImageArchived  zkf 2012107-9 update archive database   ====start
        string dbKey = "";
        if (ImageArchived == 1)
        {
            dbKey = System.Configuration.ConfigurationManager.AppSettings["ImageArchivedbKey"];
        }
        else
        {
            dbKey = System.Configuration.ConfigurationManager.AppSettings["dbKey"];
        }
        //====end

        SqlConnection sqlCnn = new SqlConnection(dbKey);
        SqlCommand sqlCmd = new SqlCommand();
        SqlDataReader sqldr = null;
        sqlCmd.Connection = sqlCnn;
        sqlCmd.CommandTimeout = 0;
        if (sourceId == 1)
        {
            sqlCmd.CommandText = "prc_DownloadDocument";
        }
        else
        {
            sqlCmd.CommandText = "prc_DownloadAttachment";
        }
        sqlCmd.CommandType = CommandType.StoredProcedure;
        sqlCmd.Parameters.Add(new SqlParameter("@ID", id));
        try
        {
            sqlCnn.Open();
            sqldr = sqlCmd.ExecuteReader();
            if (!sqldr.HasRows)
            { 
                byte[] Bytes = GetPDFBytesFromTextIsNull();
                context.Response.Buffer = true;
                context.Response.Charset = "";
                context.Response.Cache.SetCacheability(HttpCacheability.NoCache);
                context.Response.ContentType = "application/pdf";
                context.Response.BinaryWrite(Bytes);
                context.Response.Flush();
                context.Response.End();
                return;
            }
           
            while (sqldr.Read())
            {
                if (string.IsNullOrEmpty(sqldr.GetString(0)))
                {
                    byte[] Bytes = GetPDFBytesFromTextIsNull();
                    context.Response.Buffer = true;
                    context.Response.Charset = "";
                    context.Response.Cache.SetCacheability(HttpCacheability.NoCache);
                    context.Response.ContentType = "application/pdf";
                    context.Response.BinaryWrite(Bytes);
                    context.Response.Flush();
                    context.Response.End();
                    return;
                }
                else
                {
                    fileName = sqldr.GetString(0);
                    byte[] Bytes = sqldr.GetValue(1) as byte[];
                    if (Bytes == null) {
                        Bytes = GetPDFBytesFromTextIsNull();
                        context.Response.Buffer = true;
                        context.Response.Charset = "";
                        context.Response.Cache.SetCacheability(HttpCacheability.NoCache);
                        context.Response.ContentType = "application/pdf";
                        context.Response.BinaryWrite(Bytes);
                        context.Response.Flush();
                        context.Response.End();
                        return;
                    }

                    if (Bytes.Length >1024)
                    {
                        fileType = "pdf";
                        if (fileName != null)
                        {
                            int lastIndex = fileName.LastIndexOf('.');
                            fileType = fileName.Substring(lastIndex + 1);
                        }
                        fileType = fileType.ToLowerInvariant();

                        switch (fileType)
                        {
                            case "tif":
                                Bytes = GetPDFBytesFromImage(Bytes);
                                break;
                            case "tiff":
                                Bytes = GetPDFBytesFromImage(Bytes);
                                break;
                            case "jpeg":
                                Bytes = GetPDFBytesFromJPG(Bytes);
                                break;
                            case "jpg":
                                Bytes = GetPDFBytesFromJPG(Bytes);
                                break;
                            case "png":
                                Bytes = GetPDFBytesFromJPG(Bytes);
                                break;
                            case "pdf":
                                //use existing Bytes
                                break;
                            default:
                                Bytes = GetPDFBytesFromText(fileType);
                                break;
                        }

                    }
                    else
                    {
                        Bytes = GetPDFBytesFromTextIsNull();
                        context.Response.Buffer = true;
                        context.Response.Charset = "";
                        context.Response.Cache.SetCacheability(HttpCacheability.NoCache);
                        context.Response.ContentType = "application/pdf";
                        context.Response.BinaryWrite(Bytes);
                        context.Response.Flush();
                        context.Response.End();
                        return;
                    }

                    context.Response.Buffer = true;
                    context.Response.Charset = "";
                    if (context.Request.QueryString["download"] == "1")
                    {
                        context.Response.AppendHeader("Content-Disposition", "attachment; filename=" + fileName);
                    }
                    context.Response.Cache.SetCacheability(HttpCacheability.NoCache);
                    context.Response.ContentType = "application/pdf";
                    context.Response.BinaryWrite(Bytes);
                    context.Response.Flush();
                    context.Response.End();
                }

            }

        }
        catch (Exception ex)
        {
            //string errString = ex.Message;
            //string errLocation = "DisplayDocument";
            //CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session["User"], System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString()), HttpContext.Current.Request.UserHostAddress.ToString());

        }
        finally
        {
            if ((!(sqldr == null)
                        && !sqldr.IsClosed))
            {
                sqldr.Close();
            }

            sqldr.Close();
            sqlCnn.Close();
            sqlCnn.Dispose();
            sqlCmd.Dispose();
            sqlCnn = null;
            sqlCmd = null;
        }

    }

    public bool IsReusable
    {
        get
        {
            return false;
        }
    }

    private byte[] GetPDFBytesFromImage(byte[] byteArrayIn)
    {
        byte[] byteArrayOut = null;
        MemoryStream stream = null;
        MemoryStream streamOut = null;
        try
        {
            stream = new MemoryStream(byteArrayIn);
            System.Drawing.Bitmap bm = new Bitmap(stream, false);

            double dImageHeight = bm.Height;
            double dImageWidth = bm.Width;
            Document aDoc = null;
            if (dImageWidth >= dImageHeight)
            {
                aDoc = new Document(PageSize.LETTER, 10f, 10f, 10f, 10f);
            }
            else
            {
                aDoc = new Document(PageSize.LETTER, 10f, 10f, 10f, 10f);
            }

            streamOut = new System.IO.MemoryStream();
            iTextSharp.text.pdf.PdfWriter aPdfWriter = iTextSharp.text.pdf.PdfWriter.GetInstance(aDoc, streamOut);
            int iPageCount = bm.GetFrameCount(FrameDimension.Page);
            aDoc.Open();
            iTextSharp.text.pdf.PdfContentByte cb = aPdfWriter.DirectContent;
            iTextSharp.text.Image img = null;
            for (int k = 0; k < iPageCount; k++)
            {
                img = null;
                bm.SelectActiveFrame(System.Drawing.Imaging.FrameDimension.Page, k);
                //img = iTextSharp.text.Image.GetInstance(bm, System.Drawing.Imaging.ImageFormat.Bmp);
                img = iTextSharp.text.Image.GetInstance(bm, ImageFormat.Tiff);
                img.ScalePercent(72.0F / img.DpiX * 100);
                // img.ScalePercent(72.0F / img.DpiX * 100, 72.0F / img.DpiY * 100);
                // img.SetAbsolutePosition(0, 0);
                img.SetAbsolutePosition((PageSize.LETTER.Width - img.ScaledWidth) / 2, (PageSize.LETTER.Height - img.ScaledHeight) / 2);
                cb.AddImage(img);
                aDoc.NewPage();
            }
            aDoc.Close();
            aPdfWriter.Close();
            byteArrayOut = streamOut.ToArray();
        }
        catch (Exception ex)
        {
            //string errString = ex.Message;
            //string errLocation = "DisplayDocument";
        }
        finally
        {
            if (stream != null)
            {
                stream.Close();
            }
            if (streamOut != null)
            {
                streamOut.Close();
            }
        }

        return byteArrayOut;

    }

    private byte[] GetPDFBytesFromJPG(byte[] byteArrayIn)
    {
        byte[] byteArrayOut = null;
        MemoryStream stream = null;
        MemoryStream streamOut = null;
        try
        {
            // stream = new MemoryStream(byteArrayIn);
            // System.Drawing.Bitmap bm = new Bitmap(stream, false);
            iTextSharp.text.Image bm = iTextSharp.text.Image.GetInstance(byteArrayIn);

            double dImageHeight = bm.Height;
            double dImageWidth = bm.Width;
            Document aDoc = null;
            if (dImageWidth >= dImageHeight)
            {
                aDoc = new Document(PageSize.LETTER, 10f, 10f, 10f, 10f);
                aDoc.SetPageSize(iTextSharp.text.PageSize.LETTER.Rotate());
            }
            else
            {
                aDoc = new Document(PageSize.LETTER, 10f, 10f, 10f, 10f);
            }

            streamOut = new System.IO.MemoryStream();
            iTextSharp.text.pdf.PdfWriter aPdfWriter = iTextSharp.text.pdf.PdfWriter.GetInstance(aDoc, streamOut);
            aDoc.Open();
            bm.SetAbsolutePosition(0, 0);
            var scalePercent = (((aDoc.PageSize.Width / bm.Width) * 100) - 4);
            bm.ScalePercent(scalePercent);
            bm.ScaleAbsoluteHeight(aDoc.PageSize.Height);
            bm.ScaleAbsoluteWidth(aDoc.PageSize.Width);

            //Resize image depend upon your need  
            //bm.ScaleToFit(580f, 560f);
            //Give space before image  
            //bm.SpacingBefore = 30f;
            ////Give some space after the image  
            //bm.SpacingAfter = 1f;
            bm.Alignment = Element.ALIGN_CENTER;
            // aDoc.Add(paragraph); 
            aDoc.Add(bm); //add an image to the created pdf document
            aDoc.Close();
            aPdfWriter.Close();
            byteArrayOut = streamOut.ToArray();
        }
        catch (Exception ex)
        {
            //string errString = ex.Message;
            //string errLocation = "DisplayDocument";
        }
        finally
        {
            if (stream != null)
            {
                stream.Close();
            }
            if (streamOut != null)
            {
                streamOut.Close();
            }
        }

        return byteArrayOut;

    }
    private byte[] GetPDFBytesFromText(string unsupportFiltype)
    {
        byte[] byteArrayOut = null;
        MemoryStream streamOut = null;
        string contetText = string.Empty;

        try
        {

            contetText = "The file type " + unsupportFiltype + " is currently not supported by this viewer.";
            streamOut = new System.IO.MemoryStream();
            Document aDoc = new Document();
            streamOut = new System.IO.MemoryStream();
            iTextSharp.text.pdf.PdfWriter aPdfWriter = iTextSharp.text.pdf.PdfWriter.GetInstance(aDoc, streamOut);
            iTextSharp.text.Font myFont =FontFactory.GetFont("Arial,Bold", 18.0f, BaseColor.BLACK);
            aDoc.Open();
            Paragraph p1 = new Paragraph();
            p1.Alignment =Element.ALIGN_CENTER;
            p1.Font=myFont;
            p1.Add(contetText);
            aDoc.Add(p1);
            aDoc.Close();
            aPdfWriter.Close();
            byteArrayOut = streamOut.ToArray();
        }
        catch (Exception ex)
        {
            //string errString = ex.Message;
            //string errLocation = "DisplayDocument";
        }
        finally
        {
            if (streamOut != null)
            {
                streamOut.Close();
            }
        }
        return byteArrayOut;
    }

    private byte[] GetPDFBytesFromTextIsNull()
    {
        byte[] byteArrayOut = null;
        MemoryStream streamOut = null;
        string contetText = string.Empty;

        try
        {

            contetText = "This document is either empty or its file format is not supported.";
            streamOut = new System.IO.MemoryStream();
            Document aDoc = new Document();
            streamOut = new System.IO.MemoryStream();
            iTextSharp.text.pdf.PdfWriter aPdfWriter = iTextSharp.text.pdf.PdfWriter.GetInstance(aDoc, streamOut);
            iTextSharp.text.Font myFont =FontFactory.GetFont("Arial,Bold", 18.0f, BaseColor.BLACK);
            aDoc.Open();
            Paragraph p1 = new Paragraph();
            p1.Alignment =Element.ALIGN_CENTER;
            p1.Font=myFont;
            p1.Add(contetText);
            aDoc.Add(p1);
            aDoc.Close();
            aPdfWriter.Close();
            byteArrayOut = streamOut.ToArray();
        }
        catch (Exception ex)
        {
            //string errString = ex.Message;
            //string errLocation = "DisplayDocument";
        }
        finally
        {
            if (streamOut != null)
            {
                streamOut.Close();
            }
        }
        return byteArrayOut;
    }




}