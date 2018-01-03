using System;
using System.Collections;
using System.Data;
using OutSystems.HubEdition.RuntimePlatform;
using OutSystems.RuntimePublic.Db;

using iTextSharp.text.pdf;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf.security;
using iTextSharp.license;
using Org.BouncyCastle.Pkcs;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.X509;
using System.Web;

using System.Collections.Generic;

namespace OutSystems.NssFMElectronicSignatureExt {

	///
	///This class is used as an extension to an Outsystems project, allowing for a PDF to be digitally signed and marked with
	///The pertinent information.  The Signature is then tracked within the Outsystems application which insures the integrity
	///of the user and the approval.
	///
	public class CssFMElectronicSignatureExt: IssFMElectronicSignatureExt {

        /// <summary>
        /// Signs an Invoice that has been approved by the manager. Creates a visual signature at the bottom of the document as well as digitally signing it.
        /// </summary>
        /// <param name="ssPDFToSign">The PDF document of the Invoice that is being signed.</param>
        /// <param name="ssSigningCertificate">Certificate file that is used to digitally sign the document.</param>
        /// <param name="ssCertPassword">Password for the Certificate File.</param>
        /// <param name="ssSignerName">Name of the user who is signing this document.</param>
        /// <param name="ssSignatureFont">File that contains font information that will be used for the signature</param>
        /// <param name="ssSignatureFontSize">The font size for the Signature font. Different fonts are sized differently.</param>
        /// <param name="ssSigFontFileName">File name of the signature font.</param>
        /// <param name="ssSignatureID">The ID number of the electronic signature as it is being tracked in the system.</param>
        /// <param name="ssSignerUserID">The unique ID of the User who is signing the document as it is tracked in Outsystems.</param>
        /// <param name="ssSignedPDF">The PDF that has now been electronically and digitally signed.</param>
        /// <param name="ssInvoiceID">The System ID for the Invoice being signed.</param>
        /// <param name="ssDateTimeSigned">The DateTime when the Invoice was signed in users local time.</param>
        /// <param name="ssSignerIP">The IP Address of the person signing. Used in signature Location.</param>
        /// 
        public void MssSignInvoice(byte[] ssPDFToSign, byte[] ssSigningCertificate, string ssCertPassword, string ssSignerName, byte[] ssSignatureFont, decimal ssSignatureFontSize, string ssSigFontFileName, string ssSignatureID,
            string ssSignerUserID, string ssInvoiceID, DateTime ssDateTimeSigned, string ssSignerIP, out byte[] ssSignedPDF, string ssiTextKeyPath) {

            UseLicense(ssiTextKeyPath);

            byte[] tempPdf;

            addInvoiceFooter(ssPDFToSign, out tempPdf, ssSignerName, ssSignatureFont, ssSigFontFileName, 
                ssInvoiceID, ssSignerName, ssSignatureID, ssSignerUserID, ssDateTimeSigned);

            MemoryStream certificateStream = new MemoryStream(ssSigningCertificate);

            Pkcs12Store store = new Pkcs12Store(certificateStream, ssCertPassword.ToCharArray());
            String alias = "";
            ICollection<X509Certificate> chain = new List<X509Certificate>();

            foreach (string al in store.Aliases)
                if (store.IsKeyEntry(al) && store.GetKey(al).Key.IsPrivate)
                {
                    alias = al;
                    break;
                }

            AsymmetricKeyEntry pk = store.GetKey(alias);
            foreach (X509CertificateEntry c in store.GetCertificateChain(alias))
                chain.Add(c.Certificate);

            RsaPrivateCrtKeyParameters parameters = pk.Key as RsaPrivateCrtKeyParameters;

            Sign(tempPdf, chain, parameters, DigestAlgorithms.SHA512, CryptoStandard.CMS, "Approved Invoice", ssSignerIP, ssSignerName, out ssSignedPDF);
        } // MssSignInvoice


        public void UseLicense(string path)
        {
            LicenseKey.LoadLicenseFile(HttpContext.Current.Server.MapPath(path));
        }

        private static void addInvoiceFooter(byte[] oldFile, out byte[] newfile, string textToAdd, byte[] sigFont, 
            string sigFileName, string invoice_ID, string name, string sig_ID, string user_ID, DateTime date_signed)
        {
            
            //open reader
            PdfReader reader = new PdfReader(oldFile);
            MemoryStream ms = new MemoryStream();

            PdfStamper stamper = new PdfStamper(reader, ms);
			
            BaseFont bf = BaseFont.CreateFont(sigFileName, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, true, sigFont, null);
			
            Font signature_font = new Font(bf, 10);
            Font standard_font = new Font(Font.FontFamily.HELVETICA, 10);

            int pages = reader.NumberOfPages;

            for (int x = 1; x <= pages; x++)
            {
                Rectangle size = reader.GetPageSizeWithRotation(x);
                // PdfContentByte from stamper to add content to the pages over the original content
                PdfContentByte pbover = stamper.GetOverContent(x);

                PdfPTable table = new PdfPTable(5);

                // Set Cell content
                PdfPCell cell1 = new PdfPCell(new Phrase("Invoice ID: " + invoice_ID, standard_font));
                PdfPCell cell2 = new PdfPCell(new Phrase(name, signature_font));
                PdfPCell cell3 = new PdfPCell(new Phrase("SIG_ID: " + sig_ID, standard_font));
                PdfPCell cell4 = new PdfPCell(new Phrase(name + "(" + user_ID + ")", standard_font));
                PdfPCell cell5 = new PdfPCell(new Phrase(date_signed.ToString("yyyy/MM/dd"), standard_font));

                //Format Cells
                cell1.Border = Rectangle.NO_BORDER;
                cell2.Border = Rectangle.NO_BORDER;
                cell3.Border = Rectangle.NO_BORDER;
                cell4.Border = Rectangle.NO_BORDER;
                cell5.Border = Rectangle.NO_BORDER;

                cell1.UseAscender = true;
                cell2.UseAscender = true;
                cell3.UseAscender = true;
                cell4.UseAscender = true;
                cell5.UseAscender = true;

                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                cell2.HorizontalAlignment = Element.ALIGN_CENTER;
                cell3.HorizontalAlignment = Element.ALIGN_CENTER;
                cell4.HorizontalAlignment = Element.ALIGN_CENTER;
                cell5.HorizontalAlignment = Element.ALIGN_CENTER;

                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell2.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell3.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell4.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell5.VerticalAlignment = Element.ALIGN_MIDDLE;

                table.AddCell(cell1);
                table.AddCell(cell2);
                table.AddCell(cell3);
                table.AddCell(cell4);
                table.AddCell(cell5);


                float cellwidth = (size.Width - 30) / 5;

                table.SetWidthPercentage(new float[] { cellwidth, cellwidth, cellwidth, cellwidth, cellwidth }, size);

                table.WriteSelectedRows(0, -1, 15, 25, pbover);
            }

            stamper.Close();

            newfile = ms.ToArray();
        }


        public static void Sign(byte[] src, ICollection<X509Certificate> chain, ICipherParameters pk,
                         String digestAlgorithm, CryptoStandard subfilter, string reason, string location, string contact, out byte[] dest)
        {
            // Creating the reader and the stamper
            PdfReader reader = new PdfReader(src);

            MemoryStream ms = new MemoryStream();
            
            PdfStamper stamper = PdfStamper.CreateSignature(reader, ms, '\0');

            // Creating the appearance
            PdfSignatureAppearance appearance = stamper.SignatureAppearance;
            appearance.Reason = reason;
            appearance.Location = location;
            appearance.Contact = contact;
            appearance.SignatureCreator = "FlowManager";
            appearance.LocationCaption = "User IP Address";

            // Creating the signature
            IExternalSignature pks = new PrivateKeySignature(pk, digestAlgorithm);
            MakeSignature.SignDetached(appearance, pks, chain, null, null, null, 0, subfilter);

            dest = ms.ToArray();
        }

    } // CssFMElectronicSignatureExt

} // OutSystems.NssFMElectronicSignatureExt

