using System;
using System.Collections.Generic;
using System.Text;
using Nektra.Deviare2;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Drawing;

//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
// 
// NOTICE
//
// For this plugin to work, do not enclose the DeviarePlugin class in a namespace
//
// !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

public class DeviarePlugin
{

    // -------------------------------------------------------------------------

    public int OnLoad()
    {
        System.Diagnostics.Trace.WriteLine("IEPrintWatermarkHelperCS OnLoad called");
        return 0;
    }

    public void OnUnload()
    {
        System.Diagnostics.Trace.WriteLine("IEPrintWatermarkHelperCS OnUnload called");
    }

    public int OnHookAdded(INktHookInfo hookInfo, int chainIndex, string parameters)
    {
        System.Diagnostics.Trace.WriteLine("IEPrintWatermarkHelperCS OnHookAdded called [Hook: " + hookInfo.FunctionName + " @ 0x" + hookInfo.Address.ToString("X") + " / Chain:" + chainIndex.ToString() + "]");
        return 0;
    }

    //called when a hook is detached from this plugin
    public int OnHookRemoved(INktHookInfo hookInfo, int chainIndex)
    {
        System.Diagnostics.Trace.WriteLine("IEPrintWatermarkHelperCS OnHookRemoved called [Hook: " + hookInfo.FunctionName + " @ 0x" + hookInfo.Address.ToString("X") + " / Chain:" + chainIndex.ToString() + "]");
        return 0;
    }

    public string GetFunctionCallbackName(INktHookInfo hookInfo, int chainIndex)
    {
        if (hookInfo.FunctionName.Equals("XpsServices.dll!IXpsOMPageReference::SetPage", StringComparison.OrdinalIgnoreCase))
            return "OnIXpsOMPageReferenceSetPage";
        return "";
    }

    //called when a hooked function is called
    public int OnFunctionCall(INktHookInfo hookInfo, int chainIndex, INktHookCallInfoPlugin callInfo)
    {
        // Unused
        return 0;
    }


    public int OnIXpsOMPageReferenceSetPage(INktHookInfo lpHookInfo, int dwChainIndex, INktHookCallInfoPlugin lpHookCallInfoPlugin)
    {
        System.Diagnostics.Trace.WriteLine("IEPrintWaterMarkhelperCS: OnIXpsOMPageReferenceSetPage");

        try
        {
            var cMod = lpHookCallInfoPlugin.StackTrace().Module(0);

            if (cMod.Name.ToLower().EndsWith("d2d1.dll") ||
                cMod.Name.ToLower().EndsWith("mshtml.dll"))
            {
                System.Diagnostics.Trace.WriteLine(string.Format("calling module: {0}", cMod.Name.ToLower()));

                IntPtr nReg;
                if (IntPtr.Size == 4)
                {
                    nReg = lpHookCallInfoPlugin.get_Register(eNktRegister.asmRegEsp);
                    nReg = new IntPtr(nReg.ToInt32() + 8);
                    nReg = (IntPtr)Marshal.PtrToStructure(nReg, typeof(IntPtr));
                }
                else
                {
                    nReg = lpHookCallInfoPlugin.get_Register(eNktRegister.asmRegRdx);
                }

                System.Diagnostics.Trace.WriteLine(string.Format("lpPage=0x{0:x}", nReg));
                MSXPS.IXpsOMPage lpPage = (MSXPS.IXpsOMPage)Marshal.GetObjectForIUnknown(nReg);
                AddWatermark(lpPage);
            }

            lpHookCallInfoPlugin.FilterSpyMgrEvent();
        }
        catch (Exception e)
        {
            System.Diagnostics.Trace.WriteLine(string.Format("EXCEPTION: {0}") + e.Message);
        }
        return 0;
    }

    unsafe public void AddWatermark(MSXPS.IXpsOMPage page)
    {
        System.Diagnostics.Trace.WriteLine("IEPrintWaterMarkhelperCS: AddWatermark");

        try
        {
            MSXPS.XpsOMObjectFactory cXpsFactory = new MSXPS.XpsOMObjectFactoryClass();
            MSXPS.IXpsOMPage cPage = page;
            MSXPS.XPS_COLOR xpsColor;

            xpsColor = MakeXPSColor(0x80, 0, 0, 0xff);
            MSXPS.IXpsOMSolidColorBrush cXpsFillBrush = cXpsFactory.CreateSolidColorBrush(ref xpsColor, null);

            xpsColor = MakeXPSColor(0xff, 0, 0, 0);
            MSXPS.IXpsOMSolidColorBrush cXpsStrokeBrush = cXpsFactory.CreateSolidColorBrush(ref xpsColor, null);

            MSXPS.XPS_RECT xpsRect = new MSXPS.XPS_RECT()
            {
                x = 0,
                y = 0,
                width = 100,
                height = 100
            };

            MSXPS.XPS_POINT startPoint = new MSXPS.XPS_POINT()
            {
                x = xpsRect.x,
                y = xpsRect.y
            };

            MSXPS.IXpsOMGeometryFigure cRectFigure = cXpsFactory.CreateGeometryFigure(ref startPoint);

            MSXPS.XPS_SEGMENT_TYPE[] aSegmentTypes = new MSXPS.XPS_SEGMENT_TYPE[3]
            {
                MSXPS.XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE,
                MSXPS.XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE,
                MSXPS.XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE
            };

            float[] aSegmentData = new float[6] 
            {
                xpsRect.x, 
                (xpsRect.y + xpsRect.height),
                (xpsRect.x + xpsRect.width),
                (xpsRect.y + xpsRect.height), 
                (xpsRect.x + xpsRect.width), 
                xpsRect.y 
            };

            int[] aSegmentStrokes = new int[3] { 1, 1, 1 };

            cRectFigure.SetSegments(3, 6, ref aSegmentTypes[0], ref aSegmentData[0], ref aSegmentStrokes[0]);
            cRectFigure.SetIsClosed(1);
            cRectFigure.SetIsFilled(1);
            MSXPS.IXpsOMGeometry cImageRectGeometry = cXpsFactory.CreateGeometry();
            MSXPS.IXpsOMGeometryFigureCollection cGeomFigureCollection = cImageRectGeometry.GetFigures();
            cGeomFigureCollection.Append(cRectFigure);

            MSXPS.IXpsOMPath cRectPath = cXpsFactory.CreatePath();
            cRectPath.SetGeometryLocal(cImageRectGeometry);
            cRectPath.SetAccessibilityShortDescription("Red Rectangle");
            cRectPath.SetFillBrushLocal(cXpsFillBrush);
            cRectPath.SetStrokeBrushLocal(cXpsStrokeBrush);

            var cVisualColl = cPage.GetVisuals();

            cVisualColl.Append(cRectPath);

            // Create Font Resource
            //
            System.IO.DirectoryInfo dirWindowsFolder = System.IO.Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.System));

            string strFontsFolder = System.IO.Path.Combine(dirWindowsFolder.FullName, "Fonts");
            var fontStream1 = cXpsFactory.CreateReadOnlyStreamOnFile(strFontsFolder + '\\' + GetSystemFontFileName("Arial"));
            var fontStream2 = cXpsFactory.CreateReadOnlyStreamOnFile(strFontsFolder + '\\' + GetSystemFontFileName("Times New Roman"));

            var fontUri1 = cXpsFactory.CreatePartUri(string.Format("/Resources/Fonts/{0}.odttf", Guid.NewGuid()));
            var fontUri2 = cXpsFactory.CreatePartUri(string.Format("/Resources/Fonts/{0}.odttf", Guid.NewGuid()));
            MSXPS.IXpsOMFontResource font1 = cXpsFactory.CreateFontResource(fontStream1, MSXPS.XPS_FONT_EMBEDDING.XPS_FONT_EMBEDDING_NORMAL, fontUri1, 0);
            MSXPS.IXpsOMFontResource font2 = cXpsFactory.CreateFontResource(fontStream2, MSXPS.XPS_FONT_EMBEDDING.XPS_FONT_EMBEDDING_NORMAL, fontUri2, 0);

            //
            // String 1
            // 
            xpsColor = MakeXPSColor(32, 8, 0, 0);
            var cFontBrush = cXpsFactory.CreateSolidColorBrush(ref xpsColor, null);

            MSXPS.XPS_SIZE pageSize = cPage.GetPageDimensions();
            MSXPS.XPS_POINT ptZero = new MSXPS.XPS_POINT { x = 0, y = 0 };

            var cGlyphs1 = cXpsFactory.CreateGlyphs(font1);
            cGlyphs1.SetOrigin(ref ptZero);
            cGlyphs1.SetFontRenderingEmSize(128.0f);
            cGlyphs1.SetFillBrushLocal(cFontBrush);
                                    
            cGlyphs1.SetTransformLocal(MakeTransformMatrix(cXpsFactory, 
                (float)(Math.PI / 4.0),                  
                InchesToXPSUnits(1.5f), 
                pageSize.height / 1.25f ));
            
            var cGlyphEd = cGlyphs1.GetGlyphsEditor();

            cGlyphEd.SetUnicodeString("XPS Watermark");
            
            cGlyphEd.ApplyEdits();
            cVisualColl.Append(cGlyphs1);

            //
            // String 2
            // 
            xpsColor = MakeXPSColor(0x80, 0xFF, 0, 0);
            var cFontBrush2 = cXpsFactory.CreateSolidColorBrush(ref xpsColor, null);

            var cGlyphs2 = cXpsFactory.CreateGlyphs(font2);
            cGlyphs2.SetOrigin(ref ptZero);
            cGlyphs2.SetFontRenderingEmSize(64.0f);
            cGlyphs2.SetFillBrushLocal(cFontBrush2);
            cGlyphs2.SetTransformLocal(MakeTransformMatrix(cXpsFactory, (float)(Math.PI / 4.0),
                InchesToXPSUnits(2), 
                pageSize.height / 1.10f));
            
            var cGlyphEd2 = cGlyphs2.GetGlyphsEditor();

            cGlyphEd2.SetUnicodeString("Made with Nektra Deviare!");
            cGlyphEd2.ApplyEdits();
            cVisualColl.Append(cGlyphs2);


        }
        catch (Exception e)
        {
            System.Diagnostics.Trace.WriteLine(string.Format("EXCEPTION: {0}") + e.Message);
        }

    }

    unsafe private MSXPS.XPS_COLOR MakeXPSColor(byte A, byte R, byte G, byte B)
    {
        MSXPS.XPS_COLOR xpsColor;
        xpsColor.colorType = MSXPS.XPS_COLOR_TYPE.XPS_COLOR_TYPE_SRGB;

        MSXPS.__MIDL___MIDL_itf_xpsobjectmodel_0000_0000_0028* valuePtr = &xpsColor.value;
        System.Drawing.Color c = System.Drawing.Color.FromArgb(A, R, G, B);
        *((byte*)valuePtr) = c.A;
        *((byte*)valuePtr + 1) = c.R;
        *((byte*)valuePtr + 2) = c.G;
        *((byte*)valuePtr + 3) = c.B;

        return xpsColor;
    }

    private MSXPS.IXpsOMMatrixTransform MakeTransformMatrix(MSXPS.XpsOMObjectFactory omf, float angleRad, float Tx, float Ty)
    {
        MSXPS.XPS_MATRIX m;
        m.m11 = (float)Math.Cos(angleRad);
        m.m12 = (float)-Math.Sin(angleRad);
        m.m21 = (float)Math.Sin(angleRad);
        m.m22 = (float)Math.Cos(angleRad);
        m.m31 = Tx;
        m.m32 = Ty;

        return omf.CreateMatrixTransform(ref m);
    }

    private float InchesToXPSUnits(float i)
    {
        return 96.0f * i;
    }

    private static string GetSystemFontFileName(string fontName)
    {
        string fullFontName = fontName + " (TrueType)";
        RegistryKey fonts = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Fonts", false);
        if (fonts == null)
        {
            fonts = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Fonts", false);
            if (fonts == null)
            {
                System.Diagnostics.Trace.WriteLine("Cannot find font!");
                throw new Exception("Can't find font registry database.");
            }
        }
        foreach (string fntkey in fonts.GetValueNames())
        {
            if (fntkey == fullFontName)
            {
                string regFilename = fonts.GetValue(fntkey).ToString();
                System.Diagnostics.Trace.WriteLine(string.Format("Font requested={0} filename={1}", fontName, regFilename));
                return regFilename;
            }
        }

        System.Diagnostics.Trace.WriteLine(string.Format("Font requested={0} not found", fontName));
        return null;
    }
}
