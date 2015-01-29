#region PDFsharp - A .NET library for processing PDF
//
// Authors:
//   PDFsharp Team (mailto:PDFsharpSupport@pdfsharp.de)
//
// Copyright (c) 2005-2008 empira Software GmbH, Cologne (Germany)
//
// http://www.pdfsharp.com
// http://sourceforge.net/projects/pdfsharp
//
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the "Software"),
// to deal in the Software without restriction, including without limitation
// the rights to use, copy, modify, merge, publish, distribute, sublicense,
// and/or sell copies of the Software, and to permit persons to whom the
// Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included
// in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
// THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
// DEALINGS IN THE SOFTWARE.
#endregion

using System;
using System.Diagnostics;
using System.IO;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;

// By courtesy of Peter Berndts 

namespace Booklet
{
  /// <summary>
  /// This sample shows how to produce a booklet by placing
  /// two pages of an existing document on
  /// one landscape orientated page of a new document.
  /// </summary>
  class Program
  {
    [STAThread]
    static void Main(string[] args)
    {
      string[] dateien = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.pdf");
      foreach (String datei in dateien)
      {

          Console.WriteLine("Bearbeite Datei " + datei); 
          // Create the output document
          PdfDocument outputDocument = new PdfDocument();

          // Show single pages
          // (Note: one page contains two pages from the source document.
          //  If the number of pages of the source document can not be
          //  divided by 4, the first pages of the output document will
          //  each contain only one page from the source document.)
          outputDocument.PageLayout = PdfPageLayout.SinglePage;

          XGraphics gfx;
          XRect box;

          // Open the external document as XPdfForm object
          XPdfForm form = XPdfForm.FromFile(datei);
          // Determine width and height
          double ext_width = form.PixelWidth;
          double ext_height = form.PixelHeight;

          int InputPages = form.PageCount;
          int sheets = InputPages / 4; // Anzahl doppelseitig bedruckter Blätter
          if (sheets * 4 < InputPages)
              sheets += 1;
          int sheetSize = sheets % 4;
          if (sheetSize < sheets) sheetSize += 4; // Erhöhung der Seitenzahl um eins (je 2 Seiten doppelseitig)
          
          //sheetSize = 3 / sheets = 7 -> Fehler 

          int pagesLeft = InputPages;
          int vacats;
          int beginPage = 0;
          for (int sidx = 1; sidx <= sheets; sidx += sheetSize)
          {
              int allpages = 4 * sheetSize;
              beginPage = (sidx-1)*4;
              if (pagesLeft > allpages)
              {
                  pagesLeft -= allpages;
                  vacats = 0;
              }
              else
              {
                  vacats = allpages - pagesLeft;
              }
              
              Console.WriteLine("Blatt " + Convert.ToString(sidx % 4));
              for (int idx = 1; idx <= sheetSize; idx += 1)
              {
                  // Front page of a sheet:
                  // Add a new page to the output document
                  PdfPage page = outputDocument.AddPage();
                  page.Orientation = PageOrientation.Landscape;
                  page.Width = 2 * ext_width;
                  page.Height = ext_height;
                  double width = page.Width;
                  double height = page.Height;

                  gfx = XGraphics.FromPdfPage(page);

                  // Skip if left side has to remain blank
                  if (vacats > 0)
                  {
                      vacats -= 1;
                      Console.WriteLine("Leere Seite");
                  }
                  else
                  {
                      // Set page number (which is one-based) for left side
                      form.PageNumber = beginPage + allpages + 2 * (1 - idx);
                      box = new XRect(0, 0, width / 2, height);
                      // Draw the page identified by the page number like an image
                      gfx.DrawImage(form, box);
                      Console.WriteLine("Seite " + form.PageNumber.ToString());
                  }

                  // Set page number (which is one-based) for right side
                  form.PageNumber = beginPage + 2 * idx - 1;
                  box = new XRect(width / 2, 0, width / 2, height);
                  // Draw the page identified by the page number like an image
                  if ( form.PageNumber <= InputPages )
                  {
                        gfx.DrawImage(form, box);
                  }

                  if (beginPage + allpages + 1 - 2 * idx <= InputPages+vacats)
                  {
                      Console.WriteLine("Seite " + form.PageNumber.ToString());
                      // Back page of a sheet
                      page = outputDocument.AddPage();
                      page.Orientation = PageOrientation.Landscape;
                      page.Width = 2 * ext_width;
                      page.Height = ext_height;

                      gfx = XGraphics.FromPdfPage(page);
                      // Set page number (which is one-based) for left side

                      form.PageNumber = beginPage + 2 * idx;
                      box = new XRect(0, 0, width / 2, height);
                      // Draw the page identified by the page number like an image
                      if (form.PageNumber <= InputPages)
                      {
                          gfx.DrawImage(form, box);
                      }
                      Console.WriteLine("Seite " + form.PageNumber.ToString());
                      // Skip if right side has to remain blank
                      if (vacats > 0)
                      {
                          vacats -= 1;
                          Console.WriteLine("Leere Seite");
                      }
                      else
                      {
                          // Set page number (which is one-based) for right side
                          form.PageNumber = beginPage + allpages + 1 - 2 * idx;
                          box = new XRect(width / 2, 0, width / 2, height);
                          // Draw the page identified by the page number like an image
                          gfx.DrawImage(form, box);
                          Console.WriteLine("Seite " + form.PageNumber.ToString());
                      }
                  }
              }
              
          }
          // Save the document...
          
          outputDocument.Save(datei.Replace(".pdf","")+" Booklet.pdf");
      }
      // ...and start a viewer.
      //Process.Start(filename);
    }
  }
}