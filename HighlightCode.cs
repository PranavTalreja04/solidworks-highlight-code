using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;
using System.Windows;
using System.Security.AccessControl;
using System.Security.Principal;

namespace SwCSharpAddin1
{
    public enum MW_PMP_UNIT_TYPE
    {
        MW_PMP_UNIT_INTEGER = 1,
        MW_PMP_UNIT_DOUBLE = 2,
        MW_PMP_UNIT_LENGTH = 3,
        MW_PMP_UNIT_ANGLE = 4,
        MW_PMP_UNIT_DENSITY = 5,
        MW_PMP_UNIT_STRESS = 6,
        MW_PMP_UNIT_FORCE = 7,
        MW_PMP_UNIT_GRAVITY = 8,
        MW_PMP_UNIT_TIME = 9,
        MW_PMP_UNIT_FREQUENCY = 10
    }

    public class FacetVertexData
    {
        double m_x = 0.0;
        double m_y = 0.0;
        double m_z = 0.0;

        public double getx
        {
            get { return m_x; }
        }
        public void setx(double x)
        {
            m_x = x;
        }
        public double gety
        {
            get { return m_y; }
        }
        public void sety(double y)
        {
            m_y = y;
        }
        public double getz
        {
            get { return m_z; }
        }
        public void setz(double z)
        {
            m_z = z;
        }

    }

    // defining class
    public class BBoxSelectedData
    {
        IFace m_pTopFace = null;
        IFace m_pBottomFace = null;
        double m_dTopOffsetValue = 3.0;
        double m_dSideOffsetValue = 2.0;

        // parameterized constructor
        public void GetSelectData(ref IFace pTopFace, ref IFace pBottomFace, ref double dTopOffsetValue, ref double dSideOffsetValue)
        {
            pTopFace = m_pTopFace;
            pBottomFace = m_pBottomFace;
            dTopOffsetValue = m_dTopOffsetValue;
            dSideOffsetValue = m_dSideOffsetValue;
        }
        public IFace GetSelectDataTopFace()
        {
            return m_pTopFace;
        }
        public IFace GetSelectDataBottomFace()
        {
            return m_pBottomFace;
        }
        public double GetSelectDataTopOffsetValue()
        {
            return m_dTopOffsetValue;
        }
        public double GetSelectDataSideOffsetValue()
        {
            return m_dSideOffsetValue;
        }

        public void SetSelectData(IFace pTopFace, IFace pBottomFace, double dTopOffsetValue, double dSideOffsetValue)
        {
            m_pTopFace = pTopFace;
            m_pBottomFace = pBottomFace;
            m_dTopOffsetValue = dTopOffsetValue;
            m_dSideOffsetValue = dSideOffsetValue;
        }

    }

    public class UserPMPage
    {
        //Local Objects
        public IPropertyManagerPage2 swPropertyPage = null;
        PMPHandler handler = null;
        ISldWorks iSwApp = null;
        SwAddin userAddin = null;
        IPropertyManagerPageTab ppagetab1 = null;
        //IPropertyManagerPageTab ppagetab2 = null;

        #region Property Manager Page Controls
        //Groups
        IPropertyManagerPageGroup group1;
        IPropertyManagerPageGroup group2;

        //Controls
        IPropertyManagerPageTextbox textbox1;
        IPropertyManagerPageCheckbox checkbox1;
        public IPropertyManagerPageOption option1;
        public IPropertyManagerPageOption option2;
        IPropertyManagerPageOption option3;
        IPropertyManagerPageLabel lableSetTopOffset;
        IPropertyManagerPageListbox list1;

        IPropertyManagerPageSelectionbox selection1;
        IPropertyManagerPageNumberbox num1;
        IPropertyManagerPageCombobox combo1;

        IPropertyManagerPageButton button1;
        IPropertyManagerPageButton button2;
        IPropertyManagerPageButton button3;

        // IPropertyManagerPage button3;
        IPropertyManagerPageLabel lableListBoxCaption;
        public IPropertyManagerPageNumberbox numberTopOffset;
        IPropertyManagerPageLabel lableSetSideOffset;
        public IPropertyManagerPageNumberbox numberSideOffset;


        public IPropertyManagerPageTextbox textbox2;
        public IPropertyManagerPageTextbox textbox3;
        public IPropertyManagerPageTextbox textbox4;



        //Control IDs
        public const int group1ID = 0;
        public const int group2ID = 1;

        public const int textbox1ID = 2;
        public const int checkbox1ID = 3;
        public const int option1ID = 4;
        public const int option2ID = 5;
        public const int option3ID = 6;
        public const int list1ID = 7;

        public const int selection1ID = 8;
        public const int num1ID = 9;
        public const int combo1ID = 10;
        public const int tabID1 = 11;
        public const int tabID2 = 12;
        public const int buttonID1 = 13;
        public const int buttonID2 = 14;
        public const int buttonID3 = 15;
        public const int textbox2ID = 16;
        public const int textbox3ID = 17;
        public const int textbox4ID = 18;
        public const int lableSetTopOffsetID = 19;
        public const int numberTopOffsetID = 20;

        public const int lableSetSideOffsetID = 21;
        public const int numberSideOffsetID = 22;
        public const int lableListBoxCaptionID = 23;
        

        #endregion

        #region Selected Data 
        IFace pTopFace = null;
        int m_iIsTopOrBottomOptActive = 0; // If '1' set then Top Face selection is active & if '2' is set then Bottom Face selection is active. for any other number no option is set
        IFace pBottomFace = null;
        double m_dTopOffsetPageValue = 3.0;
        double m_dSideOffsetPageValue = 2.0;
        int iListItemCounter = 0;
        Dictionary<string, object> m_dictionaryBBoxSelectedData = new Dictionary<string, object>();
        List<float> m_TessTotalTriangsList = new List<float>();
        List<float> m_TessTotalNormsList = new List<float>();

        // public properties
        public IFace getTopFace
        {
            get { return pTopFace; }
        }
        public void setTopFace(IFace pLocalTopFace)
        {
            pTopFace = pLocalTopFace;
        }
        public IFace getBottomFace
        {
            get { return pBottomFace; }
        }
        public void setBottomFace(IFace pLocalBottomFace)
        {
            pBottomFace = pLocalBottomFace;
        }
        public int getIsTopOrBottomOptActive
        {
            get { return m_iIsTopOrBottomOptActive; }
        }
        public void setIsTopOrBottomOptActive(int iIsTopOrBottomOptActive)
        {
            m_iIsTopOrBottomOptActive = iIsTopOrBottomOptActive;
        }

        public double getTopOffsetPageValue
        {
            get { return m_dTopOffsetPageValue; }
        }
        public void setTopOffsetPageValue(double dTopOffsetPageValue)
        {
            m_dTopOffsetPageValue = dTopOffsetPageValue;
        }
        public double getSideOffsetPageValue
        {
            get { return m_dSideOffsetPageValue; }
        }
        public void setSideOffsetPageValue(double dSideOffsetPageValue)
        {
            m_dSideOffsetPageValue = dSideOffsetPageValue;
        }

        public void clearBBoxDictionary()
        {
            m_dictionaryBBoxSelectedData.Clear();
        }

        #endregion

        public UserPMPage(SwAddin addin)
        {
            userAddin = addin;
            if (userAddin != null)
            {
                iSwApp = (ISldWorks)userAddin.SwApp;
                CreatePropertyManagerPage();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("SwAddin not set.");
            }
        }


        protected void CreatePropertyManagerPage()
        {
            int errors = -1;
            int options = (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton |
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_CancelButton;

            handler = new PMPHandler(userAddin, this);
            swPropertyPage = (IPropertyManagerPage2)iSwApp.CreatePropertyManagerPage("Generate Stock STL dialog", options, handler, ref errors);
            if (swPropertyPage != null && errors == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
            {
                try
                {
                    AddControls();
                }
                catch (Exception e)
                {
                    iSwApp.SendMsgToUser2(e.Message, 0, 0);
                }
            }
        }


        //Controls are displayed on the page top to bottom in the order 
        //in which they are added to the object.
        protected void AddControls()
        {
            short controlType = -1;
            short align = -1;
            int options = -1;
            bool retval;

            /**
            //Add Message
            retval = swPropertyPage.SetMessage3("This is a sample message, marked yellow to signify importance.",
                                            (int)swPropertyManagerPageMessageVisibility.swImportantMessageBox,
                                            (int)swPropertyManagerPageMessageExpanded.swMessageBoxExpand,
                                            "Sample Important Caption");
            **/
            // Add PropertyManager Page Tabs
            ppagetab1 = swPropertyPage.AddTab(tabID1, "Selecting Face", "", 0);

            /**/
            //Add the groups
            options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded |
                      (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;

            group1 = (IPropertyManagerPageGroup)ppagetab1.AddGroupBox(group1ID, "To Select Face", options);

            options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Checkbox |
                      (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;


            /**/

            //Add the controls to group1

            //option1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Option;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            option1 = (IPropertyManagerPageOption)group1.AddControl(option1ID, controlType, "Top Face", align, options, "Radio Buttons");
            option1.Checked = true;
            // Top Face is active...
            setIsTopOrBottomOptActive(1);

            //option2
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Option;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            option2 = (IPropertyManagerPageOption)group1.AddControl(option2ID, controlType, "Bottom Face", align, options, "Radio Buttons");
            option2.Checked = false;

            /////////////////////////////////////////////////////////////////
            // Add Static text for Top offset
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;
            lableSetTopOffset = (IPropertyManagerPageLabel)group1.AddControl2(lableSetTopOffsetID, controlType, "Set Top offset", align, options, "Set Top offset to stock");

            // Numberbox2 for Top offset
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Numberbox;
            //align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_DoubleIndent;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            numberTopOffset = (IPropertyManagerPageNumberbox)group1.AddControl2(numberTopOffsetID, controlType, "Set Top offset", align, options, "Set Top offset to stock");
            numberTopOffset.SetRange((int)MW_PMP_UNIT_TYPE.MW_PMP_UNIT_DOUBLE, 0.0, 1000.0, 1, true);
            numberTopOffset.Value = 3.0;

            ////////////////////////////////////////////////////////////////////////////
            /// Add Static text for Side offset
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;
            lableSetSideOffset = (IPropertyManagerPageLabel)group1.AddControl2(lableSetSideOffsetID, controlType, "Set Side offset", align, options, "Set Side offset to stock");

            // Numberbox for side offset
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Numberbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            numberSideOffset = (IPropertyManagerPageNumberbox)group1.AddControl2(numberSideOffsetID, controlType, "Set Side offset", align, options, "Set Side offset to stock");
            numberSideOffset.SetRange((int)MW_PMP_UNIT_TYPE.MW_PMP_UNIT_DOUBLE, 0.0, 1000.0, 1, true);
            numberSideOffset.Value = 2.0;

            //////////////////////////////////////////////////////////////////////////////////////////////
            // Button Append Top & Bootm Pair to listbox...
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Button;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible | (int)swAddControlOptions_e.swControlOptions_SmallGapAbove;         
            button1 = (IPropertyManagerPageButton)group1.AddControl2(buttonID1, controlType, "Add Top-Bottom pair", align, options, "Add Top face and Bottom face pair");

            // Add Static text for List box selection
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;
            lableListBoxCaption = (IPropertyManagerPageLabel)group1.AddControl2(lableListBoxCaptionID, controlType, "ListBox contain selected Top face and Bottom face pair", align, options, "ListBox Top-Bottom face pair");

            // Add ListBox
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Listbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            list1 = (IPropertyManagerPageListbox)group1.AddControl2(list1ID, controlType, "ListBox Top-Bottom face pair", align, options, "ListBox contain selected Top face and Bottom face pair");
            list1.Height = 80;
            //////////////////////////////////////////////////////////////////////////////////////////////

            #region      comments section        
            /***
            // Button
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Button;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            button2 = (IPropertyManagerPageButton)group1.AddControl2(buttonID2, controlType, "Disable", align, options, "Disable the control");

            // Textbox4
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Textbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            textbox4 = (IPropertyManagerPageTextbox)group1.AddControl2(textbox4ID, controlType, "Dialogue box", align, options, " Fourth Sample Textbox text");

            // Button
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Button;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                    (int)swAddControlOptions_e.swControlOptions_Visible;

            button3 = (IPropertyManagerPageButton)group1.AddControl2(buttonID3, controlType, "Dialogue box", align, options, "Select Top Phase");
            ***/
            #endregion


        }

        public void Show()
        {
            if (swPropertyPage != null)
            {
                swPropertyPage.Show();
            }
        }

        #region PMP page action executions

        public void AddTopBottomFacePairToListBox()
        {
            // check first Top face &  Bottom face both are selected or not?
            IFace pTopLocalFace = (IFace)getTopFace;
            IFace pBottomLocalFace = (IFace)getBottomFace;
            if ((pTopLocalFace == null) || (pBottomLocalFace == null))
            {
                setTopFace(null);
                setBottomFace(null);
                return;
            }

            BBoxSelectedData objBBoxSelectedData = new BBoxSelectedData();
            objBBoxSelectedData.SetSelectData(pTopLocalFace, pBottomLocalFace, m_dTopOffsetPageValue, m_dSideOffsetPageValue);

            iListItemCounter++;
            string keyname = "Item_" + iListItemCounter.ToString()   +  "    Top-Offset = " + m_dTopOffsetPageValue.ToString()  +  "    Side-Offset = " + m_dSideOffsetPageValue.ToString();

            m_dictionaryBBoxSelectedData.Add(keyname, objBBoxSelectedData);

            list1.AddItems(keyname);

            setTopFace(null);
            setBottomFace(null);

        }

        public void OnListBoxSelectionChangeUpdateDisplay(int item)
        {
            //list1.CurrentSelection
            short shortItem = (short)item;
            string strSelectedItem = list1.ItemText[shortItem];
            //BBoxSelectedData objBBoxSelectedData = null;
            object objBBoxSelectedData = null;
            m_dictionaryBBoxSelectedData.TryGetValue(strSelectedItem, out objBBoxSelectedData);
            if (objBBoxSelectedData != null)
            {
                BBoxSelectedData pBBoxSelectedData = (BBoxSelectedData)objBBoxSelectedData;
                if (pBBoxSelectedData != null)
                {

                    setTopFace(pBBoxSelectedData.GetSelectDataTopFace());
                    setBottomFace(pBBoxSelectedData.GetSelectDataBottomFace());
                    setTopOffsetPageValue(pBBoxSelectedData.GetSelectDataTopOffsetValue());
                    setSideOffsetPageValue(pBBoxSelectedData.GetSelectDataSideOffsetValue());
                }
            }

        }

        public bool ChkIfItemSelectedInListBox()
        {
            short iSeletedItem = list1.CurrentSelection;
            bool bResult = false;
            if(-1 != iSeletedItem)
                bResult = true;

            return bResult;
        }

        public void OnRMBListBoxSelectedItemDelete()
        {
            //list1.CurrentSelection
            short iSeletedItem = list1.CurrentSelection;
            string strSelectedItem = list1.ItemText[iSeletedItem];
            if (strSelectedItem.Length > 0)
            {
                object objBBoxSelectedData = null;
                m_dictionaryBBoxSelectedData.TryGetValue(strSelectedItem, out objBBoxSelectedData);
                if (objBBoxSelectedData != null)
                {
                    BBoxSelectedData pBBoxSelectedData = (BBoxSelectedData)objBBoxSelectedData;
                    if (pBBoxSelectedData != null)
                    {
                        // Remove from the Dictionary and remove from the listBox
                        m_dictionaryBBoxSelectedData.Remove(strSelectedItem);
                    }
                }

                list1.DeleteItem(iSeletedItem);
            }
            return;
        }


        public void HighlightSelectedData(bool bHighlightStatus)
        {

            ModelDoc2 swModel = (ModelDoc2)iSwApp.ActiveDoc;
            if (swModel == null)
            {
                setTopFace(null);
                setBottomFace(null);
                return;
            }

            IFace pTopLocalFace = (IFace)getTopFace;
            if (pTopLocalFace != null)
            {
                pTopLocalFace.Highlight(bHighlightStatus);
                //int numEdges[] = pTopLocalFace.GetEdges();
            }

            IFace pBottomLocalFace = (IFace)getBottomFace;
            if (pBottomLocalFace != null)
            {
                pBottomLocalFace.Highlight(bHighlightStatus);
                //int numEdges[] = pBottomLocalFace.GetEdges();
            }

            return;
        }
        private double[] m_Transform = new double[]
        {
            1,0,0,0,
            0,1,0,0,
            0,0,1,0,
            0,0,0,1
        };

        private string BrowseFile(string defName)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "STL Files (*.stl)|*.stl";
            dlg.FileName = defName + ".stl";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                return dlg.FileName;
            }
            else
            {
                return "";
            }
        }

        private double GetCurrentPartUnitConverstionFactore()
        {
            double conversiontFactor = 1.0;

            ModelDoc2 swModel = (ModelDoc2)iSwApp.ActiveDoc;
            if (swModel == null)
            {
                setTopFace(null);
                setBottomFace(null);
                return 1;
            }

            //object objActivePartUnit = swModel.GetUnits();
            int iPartDocLengthUnit = swModel.LengthUnit;

            //swLengthUnit_e eLengthEnum = swLengthUnit_e.swMM;
            /**/
            if (iPartDocLengthUnit == (int)swLengthUnit_e.swMM)
            {
                conversiontFactor = 1000;
            }
            else if (iPartDocLengthUnit == (int)swLengthUnit_e.swCM)
            {
                conversiontFactor = 100;
            }
            else if (iPartDocLengthUnit == (int)swLengthUnit_e.swMETER)
            {
                conversiontFactor = 1;
            }
            else if (iPartDocLengthUnit == (int)swLengthUnit_e.swINCHES)
            {
                conversiontFactor = 39.3701;
            }
            else if (iPartDocLengthUnit == (int)swLengthUnit_e.swFEET)
            {
                conversiontFactor = 3.28084;
            }
            ////else if (iPartDocLengthUnit == (int)swLengthUnit_e.swFEETINCHES)
            ////{
            ////    conversiontFactor = 1;
            ////}

            return conversiontFactor;
        }


        private void ExportToStl(string filePath, float[] tessTriangs, float[] tessNorms, double[] transformMatrix)
        {

            // As we are getting inputs values always in meter..
            // //then while creation of slt we will create with respective to current part file units...
            // Part units can be MM or Meter or Incheas or CM...
            // Accordingly we will change the inputs meter values to respective part units & we will generate stl position accoridnly.. 

            // So create a method here which will return the part converstion factore...
            double dLengthUnitConverstionFactor = GetCurrentPartUnitConverstionFactore();

            IMathUtility mathUtils = iSwApp.IGetMathUtility();
            IMathTransform transform = (mathUtils.CreateTransform(transformMatrix) as IMathTransform).IInverse();

            using (FileStream fileStream = File.Create(filePath))
            {
                using (BinaryWriter writer = new BinaryWriter(fileStream))
                {
                    byte[] header = new byte[80];

                    writer.Write(header);

                    uint triangsCount = (uint)tessTriangs.Length / 9;
                    writer.Write(triangsCount);

                    for (uint i = 0; i < triangsCount; i++)
                    {
                        //Debug.Print("Start Facet = " + i);
                        ////float normalX = tessNorms[i * 9];
                        ////float normalY = tessNorms[i * 9 + 1];
                        ////float normalZ = tessNorms[i * 9 + 2];
                        float normalX = tessNorms[i * 3];
                        float normalY = tessNorms[i * 3 + 1];
                        float normalZ = tessNorms[i * 3 + 2]; 

                        IMathVector mathVec = mathUtils.CreateVector(
                            new double[] { normalX, normalY, normalZ }) as IMathVector;

                        //mathVec = mathVec.MultiplyTransform(transform) as IMathVector;

                        double[] vec = mathVec.ArrayData as double[];
                        //Debug.Print("vec = " + "(" + vec[0] + ", " + vec[1] + ", " + vec[2] + ")");

                        vec[0] = vec[0] * dLengthUnitConverstionFactor;
                        vec[1] = vec[1] * dLengthUnitConverstionFactor;
                        vec[2] = vec[2] * dLengthUnitConverstionFactor;
                        writer.Write((float)vec[0]);
                        writer.Write((float)vec[1]);
                        writer.Write((float)vec[2]);

                        #region Commented
                        /***
                        for (uint j = 0; j < 3; j++)
                        {
                            float vertX = tessTriangs[i * 9 + j * 3];
                            float vertY = tessTriangs[i * 9 + j * 3 + 1];
                            float vertZ = tessTriangs[i * 9 + j * 3 + 2];

                            IMathPoint mathPt = mathUtils.CreatePoint(
                                new double[] { vertX, vertY, vertZ }) as IMathPoint;

                            mathPt = mathPt.MultiplyTransform(transform) as IMathPoint;

                            double[] pt = mathPt.ArrayData as double[];

                            writer.Write((float)pt[0]);
                            writer.Write((float)pt[1]);
                            writer.Write((float)pt[2]);
                        }
                        **/
                        #endregion 

                        /**/
                        //////////////////////////////////////////////////////////
                        float vert1X = tessTriangs[i * 9 ];
                        float vert1Y = tessTriangs[i * 9 + 1];
                        float vert1Z = tessTriangs[i * 9 + 2];

                        IMathPoint mathPt1 = mathUtils.CreatePoint(
                            new double[] { vert1X, vert1Y, vert1Z }) as IMathPoint;

                        //mathPt1 = mathPt1.MultiplyTransform(transform) as IMathPoint;

                        double[] pt1 = mathPt1.ArrayData as double[];
                        //Debug.Print("pt1 = " + "(" + pt1[0] + ", " + pt1[1] + ", " + pt1[2] + ")");
                        pt1[0] = pt1[0] * dLengthUnitConverstionFactor;
                        pt1[1] = pt1[1] * dLengthUnitConverstionFactor;
                        pt1[2] = pt1[2] * dLengthUnitConverstionFactor;
                        writer.Write((float)pt1[0]);
                        writer.Write((float)pt1[1]);
                        writer.Write((float)pt1[2]);
                        /////////////////////////////////////
                        ///                            
                        float vert2X = tessTriangs[i * 9 + 3];
                        float vert2Y = tessTriangs[i * 9 + 4];
                        float vert2Z = tessTriangs[i * 9 + 5];

                        IMathPoint mathPt2 = mathUtils.CreatePoint(
                            new double[] { vert2X, vert2Y, vert2Z }) as IMathPoint;

                        //mathPt2 = mathPt2.MultiplyTransform(transform) as IMathPoint;

                        double[] pt2 = mathPt2.ArrayData as double[];
                        //Debug.Print("pt2 = " + "(" + pt2[0] + ", " + pt2[1] + ", " + pt2[2] + ")");
                        pt2[0] = pt2[0] * dLengthUnitConverstionFactor;
                        pt2[1] = pt2[1] * dLengthUnitConverstionFactor;
                        pt2[2] = pt2[2] * dLengthUnitConverstionFactor;
                        writer.Write((float)pt2[0]);
                        writer.Write((float)pt2[1]);
                        writer.Write((float)pt2[2]);
                        /////////////////////////////////////////////////////////////
                        float vert3X = tessTriangs[i * 9 + 6];
                        float vert3Y = tessTriangs[i * 9 + 7];
                        float vert3Z = tessTriangs[i * 9 + 8];

                        IMathPoint mathPt3 = mathUtils.CreatePoint(
                            new double[] { vert3X, vert3Y, vert3Z }) as IMathPoint;

                        //mathPt3 = mathPt3.MultiplyTransform(transform) as IMathPoint;

                        double[] pt3 = mathPt3.ArrayData as double[];
                        //Debug.Print("pt3 = " + "(" + pt3[0] + ", " + pt3[1] + ", " + pt3[2] + ")");
                        pt3[0] = pt3[0] * dLengthUnitConverstionFactor;
                        pt3[1] = pt3[1] * dLengthUnitConverstionFactor;
                        pt3[2] = pt3[2] * dLengthUnitConverstionFactor;
                        writer.Write((float)pt3[0]);
                        writer.Write((float)pt3[1]);
                        writer.Write((float)pt3[2]);
                        ///////////////////////////////////////////
                        ///
                        /**/
                        ushort atts = 0;
                        writer.Write(atts);
                    }
                }
            }

            string strSuccessfulSTLMsg = "Successfully created stl at " + filePath;
            System.Windows.Forms.MessageBox.Show(strSuccessfulSTLMsg);
        }


        private void CalculateShortestDist(double dXPlanePt, double dYPlanePt, double dZPlanePt,
                                           double diPlaneNormal, double djPlaneNormal, double dkPlaneNormal,
                                           double dPtToProjx, double dPtToProjy, double dPtToProjz, ref double dShortestDistance, ref double dRefDistance)
        {
            //Project top face position on Bottom face plane.
            // calculate D
            double D = -(dXPlanePt * diPlaneNormal +
                        dYPlanePt * djPlaneNormal +
                        dZPlanePt * dkPlaneNormal);

            dRefDistance = (dPtToProjx * diPlaneNormal +
                        dPtToProjy * djPlaneNormal +
                        dPtToProjz * dkPlaneNormal + D);

            // project the pt onto plane
            //pt.set_x(TopFaceBoxDblArray[0] - dBotFaceNormalVecData[0] * d);
            // pt.set_y(TopFaceBoxDblArray[1] - dBotFaceNormalVecData[1] * d);
            // pt.set_z(TopFaceBoxDblArray[2] - dBotFaceNormalVecData[2] * d);

            // check if dRefDistance == 0
            dShortestDistance = (dPtToProjx * diPlaneNormal +
                                 dPtToProjy * djPlaneNormal +
                                 dPtToProjz * dkPlaneNormal + D);
            return;
        }

        private void CalcFarmostPt(double dXPlanePt, double dYPlanePt, double dZPlanePt,
                                   double diPlaneNormal, double djPlaneNormal, double dkPlaneNormal,
                                   double[] dPtToProjArr, ref int iGetFarmostPt, ref double dShortestDistance, ref double dRefDistance)
        {

            // for first point of Top face :- 
            double dFirstPtShortestDistance = 0.0;
            double dFirstPtRefDistance = 0.0;
            CalculateShortestDist(dXPlanePt, dYPlanePt, dZPlanePt,
                                  diPlaneNormal, djPlaneNormal, dkPlaneNormal,
                                  dPtToProjArr[0], dPtToProjArr[1], dPtToProjArr[2], ref dFirstPtShortestDistance, ref dFirstPtRefDistance);

            // for second point of Top face :- 
            double dSecondPtShortestDistance = 0.0;
            double dSecondPtRefDistance = 0.0;
            CalculateShortestDist(dXPlanePt, dYPlanePt, dZPlanePt,
                                   diPlaneNormal, djPlaneNormal, dkPlaneNormal,
                                   dPtToProjArr[3], dPtToProjArr[4], dPtToProjArr[5], ref dSecondPtShortestDistance, ref dSecondPtRefDistance);

            if (dFirstPtShortestDistance > dSecondPtShortestDistance)
            {
                iGetFarmostPt = 1;
                dShortestDistance = dFirstPtShortestDistance;
                dRefDistance = dFirstPtRefDistance;
            }
            else
            {
                iGetFarmostPt = 2;
                dShortestDistance = dSecondPtShortestDistance;
                dRefDistance = dSecondPtRefDistance;
            }

            return;
        }

        private bool CktToMakeXAxisValuesReverse(double diNormalValues, double djNormalValues, double dkNormalValues)
        {
            bool bToMakeXAxisValuesReverse = false;

            IMathUtility mathUtil = iSwApp.IGetMathUtility();

            IMathVector vecFaceNormal = mathUtil.CreateVector(new double[] { diNormalValues, djNormalValues, dkNormalValues }) as IMathVector;
            if (vecFaceNormal == null)
                return false;

            vecFaceNormal = vecFaceNormal.Normalise();

            IMathVector vecZAxis = mathUtil.CreateVector(new double[] { 0, 0, 1.0 }) as IMathVector;
            if (vecZAxis == null)
                return false;

            // double CosAngle = UVect1 % UVect2;
            double dotProd = diNormalValues * 0 +
                             djNormalValues * 0 +
                             dkNormalValues * 1;

            if (dotProd > 1.0)
                dotProd = 1.0;
            else if (dotProd < -1.0)
                dotProd = -1.0;



            // FRVector CrossProd = UVect1 * UVect2;
            // double diCrossProd = v1.comp[1] * v2.comp[2] - v1.comp[2] * v2.comp[1];
            double diCrossProd = djNormalValues * 1 - dkNormalValues * 0;
            //double djCrossProd = v1.comp[2] * v2.comp[0] - v1.comp[0] * v2.comp[2];
            double djCrossProd = dkNormalValues * 0 - diNormalValues * 1;
            //double dkCrossProd = v1.comp[0] * v2.comp[1] - v1.comp[1] * v2.comp[0];
            double dkCrossProd = diNormalValues * 0 - djNormalValues * 0;

            //double SinAngle = CrossProd.len();
            double SinAngle = Math.Sqrt(diCrossProd * diCrossProd +
                                        djCrossProd * djCrossProd +
                                        dkCrossProd * dkCrossProd);

            //IncludedAngle = atan2(SinAngle, CosAngle);
            double IncludedAngle = Math.Atan2(SinAngle, dotProd);

            //double IncludedAngleAbs = fabs(IncludedAngle);
            double IncludedAngleAbs = Math.Abs(IncludedAngle);

            //if (IncludedAngleAbs < RESNOR)
            if (IncludedAngleAbs < 0.0001)
            {
                IncludedAngle = 0.0;
            }

            if (IncludedAngle < 0.0)
            {
                IncludedAngle += 2 * Math.PI;
            }

            // double dAngleValueInPIRadian = vecZAxis.AngleTo(vecFaceNormal);   // we will get value in range 0 to PI [3.14 i.e. 180deg] 
            double dAngleValueInPIRadian = IncludedAngle;
            double dAngleValueInPIDeg = dAngleValueInPIRadian * (180 / Math.PI);

            if (dAngleValueInPIDeg > 90)   // We need to add check
                bToMakeXAxisValuesReverse = true;

            return bToMakeXAxisValuesReverse;
        }


        /// <summary>
        /// Test a directory for create file access permissions
        /// </summary>
        /// <param name="DirectoryPath">Full path to directory </param>
        /// <param name="AccessRight">File System right tested</param>
        /// <returns>State [bool]</returns>
        public static bool DirectoryHasPermission(string DirectoryPath, FileSystemRights AccessRight)
        {
            if (string.IsNullOrEmpty(DirectoryPath)) return false;

            try
            {
                AuthorizationRuleCollection rules = Directory.GetAccessControl(DirectoryPath).GetAccessRules(true, true, typeof(System.Security.Principal.SecurityIdentifier));
                WindowsIdentity identity = WindowsIdentity.GetCurrent();

                foreach (FileSystemAccessRule rule in rules)
                {
                    if (identity.Groups.Contains(rule.IdentityReference))
                    {
                        if ((AccessRight & rule.FileSystemRights) == AccessRight)
                        {
                            if (rule.AccessControlType == AccessControlType.Allow)
                                return true;
                        }
                    }
                }
            }
            catch { }
            return false;
        }


        public void GenerateSTLFileFromBBoxDictionary()
        {
            m_TessTotalTriangsList.Clear();
            m_TessTotalNormsList.Clear();

            // First the path of current active document/part...

            // ..If that part file path contain part file name ... C:\temp\abc.prt 
            // Remove last portion of string till . from reverse... so you should get C:\temp\abc. + stl
            // so finally you should get C:\temp\abc.stl 
            // Not check if the path "C:\temp\" is have write access or not..
            // If you havewrite access then create STL with above path... else run code from Browser...

            ModelDoc2 swModel = (ModelDoc2)iSwApp.ActiveDoc;
            if (swModel == null)
            {
                setTopFace(null);
                setBottomFace(null);
                return;
            }

            string newFilePath = "";

            string pathName = swModel.GetPathName();  // C:\temp\abc.prt 
            Debug.Print("Path and name of open document: " + pathName);

            string newFilePathSTL = Path.ChangeExtension(pathName, "stl");
            //string strFullPath = Path.GetFullPath(pathName);
            string strPrtFileName = Path.GetFileName(pathName);
            string strDirectoryName = Path.GetDirectoryName(pathName);
            string strPartNameWithoutExt = Path.GetFileNameWithoutExtension(pathName);
            //Path.
            //FileSystemRights AccessRight;

            bool bIsWriteAccess = false;
            bIsWriteAccess = DirectoryHasPermission(strDirectoryName, FileSystemRights.CreateFiles);
           if(true == bIsWriteAccess)
            {
                string strCreateFullFilePath = strDirectoryName + "\\" + strPartNameWithoutExt + "_Stock.stl";
                newFilePath = strCreateFullFilePath;
            }
            else
            {
                // First ask for stl file path to be created 
                string fileNameBase = "1";
                newFilePath = BrowseFile(Path.GetFileNameWithoutExtension(fileNameBase));
                if (string.IsNullOrEmpty(newFilePath))
                {
                    // Set both Bottom & Top face to null, so that to be ready for next selection set.
                    System.Windows.Forms.MessageBox.Show("Seleted stl path is empty.");
                    setTopFace(null);
                    setBottomFace(null);
                    return;
                }
            }

            foreach (var BBoxSelItem in m_dictionaryBBoxSelectedData)
            {
                object objBBoxSelectedData = BBoxSelItem.Value;
                if (objBBoxSelectedData == null)
                    continue;

                BBoxSelectedData pBBoxSelectedData = (BBoxSelectedData)objBBoxSelectedData;
                if (pBBoxSelectedData == null)
                    continue;

                GenerateSTLFile(pBBoxSelectedData);
            }

            /**/
            //Export tessalation data to stl file...
            float[] tessTotalTriangs = m_TessTotalTriangsList.ToArray();
            float[] tessTotalNorms = m_TessTotalNormsList.ToArray();
            ExportToStl(newFilePath, tessTotalTriangs, tessTotalNorms, m_Transform);
            m_TessTotalTriangsList.Clear();
            m_TessTotalNormsList.Clear();
            /**/

            return;
        }

        public void CreateFacetsForSideFaceUsingDiagonal(FacetVertexData v1, FacetVertexData v2, FacetVertexData v3, FacetVertexData v4)
        {
            #region Calculate first triangle....
            // Calculate normal for first facet
            m_TessTotalTriangsList.Add((float)v1.getx); m_TessTotalTriangsList.Add((float)v1.gety); m_TessTotalTriangsList.Add((float)v1.getz);
            m_TessTotalTriangsList.Add((float)v2.getx); m_TessTotalTriangsList.Add((float)v2.gety); m_TessTotalTriangsList.Add((float)v2.getz);
            m_TessTotalTriangsList.Add((float)v3.getx); m_TessTotalTriangsList.Add((float)v3.gety); m_TessTotalTriangsList.Add((float)v3.getz);

            IMathUtility mathUtil = iSwApp.IGetMathUtility();
            IMathVector vec11FacetNormal = mathUtil.CreateVector(new double[] { (v2.getx - v1.getx), (v2.gety - v1.gety), (v2.getz - v1.getz) }) as IMathVector;
            if (vec11FacetNormal == null)
                return;
            vec11FacetNormal = vec11FacetNormal.Normalise();
            IMathVector vec12FacetNormal = mathUtil.CreateVector(new double[] { (v3.getx - v2.getx), (v3.gety - v2.gety), (v3.getz - v2.getz) }) as IMathVector;
            if (vec12FacetNormal == null)
                return;
            vec12FacetNormal = vec12FacetNormal.Normalise();

            object objVecFirstFacetNormal = vec11FacetNormal.Cross(vec12FacetNormal);

            IMathVector vecFirstFacetNormal = objVecFirstFacetNormal as IMathVector;
            if (vecFirstFacetNormal == null)
                return;
            vecFirstFacetNormal = vecFirstFacetNormal.Normalise();
            object objFirstFacetNormalVecData = vecFirstFacetNormal.ArrayData;
            double[] dFirstFacetNormalVecData = new double[4];
            dFirstFacetNormalVecData = (double[])objFirstFacetNormalVecData;
            m_TessTotalNormsList.Add((float)dFirstFacetNormalVecData[0]); m_TessTotalNormsList.Add((float)dFirstFacetNormalVecData[1]); m_TessTotalNormsList.Add((float)dFirstFacetNormalVecData[2]);
            #endregion

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            #region Calculate 2nd triangle....
            // Calculate normal of 2nd facet
            m_TessTotalTriangsList.Add((float)v2.getx); m_TessTotalTriangsList.Add((float)v2.gety); m_TessTotalTriangsList.Add((float)v2.getz);
            m_TessTotalTriangsList.Add((float)v4.getx); m_TessTotalTriangsList.Add((float)v4.gety); m_TessTotalTriangsList.Add((float)v4.getz);
            m_TessTotalTriangsList.Add((float)v3.getx); m_TessTotalTriangsList.Add((float)v3.gety); m_TessTotalTriangsList.Add((float)v3.getz);

            IMathVector vec21FacetNormal = mathUtil.CreateVector(new double[] { (v4.getx - v2.getx), (v4.gety - v2.gety), (v4.getz - v2.getz) }) as IMathVector;
            if (vec21FacetNormal == null)
                return;
            vec21FacetNormal = vec21FacetNormal.Normalise();
            IMathVector vec22FacetNormal = mathUtil.CreateVector(new double[] { (v3.getx - v4.getx), (v3.gety - v4.gety), (v3.getz - v4.getz) }) as IMathVector;
            if (vec22FacetNormal == null)
                return;
            vec22FacetNormal = vec22FacetNormal.Normalise();

            object objVecSecondFacetNormal = vec21FacetNormal.Cross(vec22FacetNormal);
            IMathVector vecSecondFacetNormal = objVecSecondFacetNormal as IMathVector;
            if (vecSecondFacetNormal == null)
                return;
            vecSecondFacetNormal = vecSecondFacetNormal.Normalise();
            object objSecondFacetNormalVecData = vecSecondFacetNormal.ArrayData;
            double[] dSecondFacetNormalVecData = new double[4];
            dSecondFacetNormalVecData = (double[])objSecondFacetNormalVecData;

            m_TessTotalNormsList.Add((float)dSecondFacetNormalVecData[0]); m_TessTotalNormsList.Add((float)dSecondFacetNormalVecData[1]); m_TessTotalNormsList.Add((float)dSecondFacetNormalVecData[2]);
            #endregion

            return;
        }

        public void CreateFacetForRectangleExtrude(double dXLowerOfBottomFace, double dYLowerOfBottomFace, double dZLowerOfBottomFace,
                                                    double dXUpperOfBottomFace, double dYUpperOfBottomFace, double dZOfTopFace)
        {

            //    V3_______/________V4 _______________V6 
            //    |\       \       |\                 |          
            //    | \              |  \               |
            //    |   \     F2     |    \      F4     |
            //    |     \ _       /|\      \          |
            //    |      |\        |          \       |
            //    |   F1    \      |    F3       \    |
            //    |           \    |               \  |
            //    |_______\______\_|_________________\|
            //    V1      /        V2                 V5

            // So each facet triangulation start using Righthand rule for CCW direction required for faceting
            //  F1 = V1,V2,V3, F2 = V2,V4,V3 , F3= V2,V5,V4, F4=V5,V6,V4.  F5= V5,V7,V6, F6=V.........
            //
            // Pass two diagonal point of one side face of rectangle extrude & append its two triangle data in list with its normal details
            // For whole cube type there will be total 6 side faces..
            #region Side 1 of rectangular cube:-
            FacetVertexData vt1 = new FacetVertexData();
            vt1.setx(dXLowerOfBottomFace);
            vt1.sety(dYLowerOfBottomFace);
            vt1.setz(dZLowerOfBottomFace);
            FacetVertexData vt2 = new FacetVertexData();
            vt2.setx(dXUpperOfBottomFace);
            vt2.sety(dYLowerOfBottomFace);
            vt2.setz(dZLowerOfBottomFace);
            FacetVertexData vt3 = new FacetVertexData();
            vt3.setx(dXLowerOfBottomFace);
            vt3.sety(dYLowerOfBottomFace);
            vt3.setz(dZOfTopFace);
            FacetVertexData vt4 = new FacetVertexData();
            vt4.setx(dXUpperOfBottomFace);
            vt4.sety(dYLowerOfBottomFace);
            vt4.setz(dZOfTopFace);
            CreateFacetsForSideFaceUsingDiagonal(vt1, vt2, vt3, vt4);
            #endregion

            //////////////////////////////////////////////////////////////
            #region Side 2 of rectangular cube:-
            vt1.setx(dXUpperOfBottomFace);
            vt1.sety(dYLowerOfBottomFace);
            vt1.setz(dZLowerOfBottomFace);

            vt2.setx(dXUpperOfBottomFace);
            vt2.sety(dYUpperOfBottomFace);
            vt2.setz(dZLowerOfBottomFace);

            vt3.setx(dXUpperOfBottomFace);
            vt3.sety(dYLowerOfBottomFace);
            vt3.setz(dZOfTopFace);

            vt4.setx(dXUpperOfBottomFace);
            vt4.sety(dYUpperOfBottomFace);
            vt4.setz(dZOfTopFace);
            CreateFacetsForSideFaceUsingDiagonal(vt1, vt2, vt3, vt4);
            #endregion

            //////////////////////////////////////////////////////////////
            #region Side 3 of rectangular cube:-
            vt1.setx(dXUpperOfBottomFace);
            vt1.sety(dYUpperOfBottomFace);
            vt1.setz(dZLowerOfBottomFace);

            vt2.setx(dXLowerOfBottomFace);
            vt2.sety(dYUpperOfBottomFace);
            vt2.setz(dZLowerOfBottomFace);

            vt3.setx(dXUpperOfBottomFace);
            vt3.sety(dYUpperOfBottomFace);
            vt3.setz(dZOfTopFace);

            vt4.setx(dXLowerOfBottomFace);
            vt4.sety(dYUpperOfBottomFace);
            vt4.setz(dZOfTopFace);
            CreateFacetsForSideFaceUsingDiagonal(vt1, vt2, vt3, vt4);
            #endregion

            //////////////////////////////////////////////////////////////
            #region Side 4 of rectangular cube:-
            vt1.setx(dXLowerOfBottomFace);
            vt1.sety(dYUpperOfBottomFace);
            vt1.setz(dZLowerOfBottomFace);

            vt2.setx(dXLowerOfBottomFace);
            vt2.sety(dYLowerOfBottomFace);
            vt2.setz(dZLowerOfBottomFace);

            vt3.setx(dXLowerOfBottomFace);
            vt3.sety(dYUpperOfBottomFace);
            vt3.setz(dZOfTopFace);

            vt4.setx(dXLowerOfBottomFace);
            vt4.sety(dYLowerOfBottomFace);
            vt4.setz(dZOfTopFace);
            CreateFacetsForSideFaceUsingDiagonal(vt1, vt2, vt3, vt4);
            #endregion

            #region Now Side 5 bottom face of cube shape.....
            vt1.setx(dXLowerOfBottomFace);
            vt1.sety(dYLowerOfBottomFace);
            vt1.setz(dZLowerOfBottomFace);

            vt2.setx(dXLowerOfBottomFace);
            vt2.sety(dYUpperOfBottomFace);
            vt2.setz(dZLowerOfBottomFace);

            vt3.setx(dXUpperOfBottomFace);
            vt3.sety(dYLowerOfBottomFace);
            vt3.setz(dZLowerOfBottomFace);

            vt4.setx(dXUpperOfBottomFace);
            vt4.sety(dYUpperOfBottomFace);
            vt4.setz(dZLowerOfBottomFace);
            CreateFacetsForSideFaceUsingDiagonal(vt1, vt2, vt3, vt4);
            #endregion

            #region Side 6 Top face of cube shape.....
            vt1.setx(dXLowerOfBottomFace);
            vt1.sety(dYLowerOfBottomFace);
            vt1.setz(dZOfTopFace);

            vt2.setx(dXUpperOfBottomFace);
            vt2.sety(dYLowerOfBottomFace);
            vt2.setz(dZOfTopFace);

            vt3.setx(dXLowerOfBottomFace);
            vt3.sety(dYUpperOfBottomFace);
            vt3.setz(dZOfTopFace);

            vt4.setx(dXUpperOfBottomFace);
            vt4.sety(dYUpperOfBottomFace);
            vt4.setz(dZOfTopFace);
            CreateFacetsForSideFaceUsingDiagonal(vt1, vt2, vt3, vt4);
            #endregion

            return;
        }

        public void GenerateSTLFile(BBoxSelectedData pBBoxSelectedIteamDatac)
        {

            // check first Top Face is selected or not?
            IFace pTopLocalFace = (IFace)pBBoxSelectedIteamDatac.GetSelectDataTopFace();
            IFace pBottomLocalFace = (IFace)pBBoxSelectedIteamDatac.GetSelectDataBottomFace();
            
            double dTopOffsetValue = pBBoxSelectedIteamDatac.GetSelectDataTopOffsetValue();
            double dSideOffsetValue = pBBoxSelectedIteamDatac.GetSelectDataSideOffsetValue();

            if ((pTopLocalFace == null) || (pBottomLocalFace == null))
            {
                setTopFace(null);
                setBottomFace(null);
                return;
            }

            IFace pBottomLocalFaceTest = pBottomLocalFace as IFace;
            object objBottomFaceNormal = null;
            objBottomFaceNormal = pBottomLocalFace.Normal;  //if Face is not Planer then we will get (0,0,0)
            double[] dBottomFaceNormalValues = new double[4];
            dBottomFaceNormalValues = (double[])objBottomFaceNormal;

            //Debug.Print("Face  Normal = " + "(" + dBottomFaceNormalValues[0] + ", " + dBottomFaceNormalValues[1] + ", " + dBottomFaceNormalValues[2] + ")");
            if ((dBottomFaceNormalValues[0] == 0) && (dBottomFaceNormalValues[1] == 0) && (dBottomFaceNormalValues[2] == 0))
            {
                // Selected bottom face is not planer
                setTopFace(null);
                setBottomFace(null);
                return;
            }

            //bool bToRevereXAxisValue = CktToMakeXAxisValuesReverse(dBottomFaceNormalValues[0], dBottomFaceNormalValues[1], dBottomFaceNormalValues[2]);

            ////////////////////////////////////////////////////////////////////////
            ///Bottom face calculation......
            //int iFaceSelID = face.GetFaceId();
            object BottomFaceBoxArray = null;
            double[] BottomFaceBoxDblArray = new double[7];

            BottomFaceBoxArray = (object)pBottomLocalFace.GetBox();
            BottomFaceBoxDblArray = (double[])BottomFaceBoxArray;
            //Debug.Print("Face =");
            //Debug.Print("  Pt1 = " + "(" + BottomFaceBoxDblArray[0] * 1000.0 + ", " + BottomFaceBoxDblArray[1] * 1000.0 + ", " + BottomFaceBoxDblArray[2] * 1000.0 + ") mm");
            //Debug.Print("  Pt2 = " + "(" + BottomFaceBoxDblArray[3] * 1000.0 + ", " + BottomFaceBoxDblArray[4] * 1000.0 + ", " + BottomFaceBoxDblArray[5] * 1000.0 + ") mm");

            //Write similar above GetBox code fro TopFace....
            ////////////////////////////////////////////////////////////////////////////
            object TopFaceBoxArray = null;
            double[] TopFaceBoxDblArray = new double[7];
            TopFaceBoxArray = (object)pTopLocalFace.GetBox();
            TopFaceBoxDblArray = (double[])TopFaceBoxArray;
            //Debug.Print("Top Face =");
            //Debug.Print("  Pt1 = " + "(" + TopFaceBoxDblArray[0] * 1000.0 + ", " + TopFaceBoxDblArray[1] * 1000.0 + ", " + TopFaceBoxDblArray[2] * 1000.0 + ") mm");
            //Debug.Print("  Pt2 = " + "(" + TopFaceBoxDblArray[3] * 1000.0 + ", " + TopFaceBoxDblArray[4] * 1000.0 + ", " + TopFaceBoxDblArray[5] * 1000.0 + ") mm");

            ///////////////////////////////////////////////////////////////////////////
            //Take there.. X-Y from Bottomface & Z from TopFace....
            //double dXLowerOfBottomFace = BottomFaceBoxDblArray[0] - 0.002;
            double dXLowerOfBottomFace = BottomFaceBoxDblArray[0] - dSideOffsetValue/1000;
            
            // This is required for sketch to draw, because when Bottom Face Normal is (0,0,-1) then the sketch rectangle is drawing in opposite side.
            //if (true == bToRevereXAxisValue)
            //    dXLowerOfBottomFace = -dXLowerOfBottomFace;
            //double dYLowerOfBottomFace = BottomFaceBoxDblArray[1] - 0.002;
            double dYLowerOfBottomFace = BottomFaceBoxDblArray[1] - dSideOffsetValue/1000;
            double dZLowerOfBottomFace = BottomFaceBoxDblArray[2];

            //double dXUpperOfBottomFace = BottomFaceBoxDblArray[3] + 0.002;
            double dXUpperOfBottomFace = BottomFaceBoxDblArray[3] + dSideOffsetValue/1000;
            // This is required for sketch to draw, because when Bottom Face Normal is (0,0,-1) then the sketch rectangle is drawing in opposite side.
            //if (true == bToRevereXAxisValue)
            //    dXUpperOfBottomFace = -dXUpperOfBottomFace;
            //double dYUpperOfBottomFace = BottomFaceBoxDblArray[4] + 0.002;
            double dYUpperOfBottomFace = BottomFaceBoxDblArray[4] + dSideOffsetValue/1000;
            double dZUpperOfBottomFace = BottomFaceBoxDblArray[5];

            //double dZOfTopFace = TopFaceBoxDblArray[2] + 0.003;
            //////////////////////////////////////////////////////////////////////////////////
            //Calculate the shortest distance between Bottom planer face and Top face :-
            // consider both bounding box position of Top face & take fartherest position to calculate the distance from these two BBox position.
            IMathUtility mathUtil = iSwApp.IGetMathUtility();

            IMathVector dirVec = mathUtil.CreateVector(objBottomFaceNormal) as IMathVector;
            double dBeforeNormalization = dirVec.GetLength();
            dirVec = dirVec.Normalise();
            double dAfterNormalization = dirVec.GetLength();

            object objBotFaceNormalVecData = dirVec.ArrayData;
            double[] dBotFaceNormalVecData = new double[4];
            dBotFaceNormalVecData = (double[])objBotFaceNormalVecData;

            int iGetFarmostPt = 1;
            double dShortestDistance = 0.0;
            double dRefDistance = 0.0;
            CalcFarmostPt(BottomFaceBoxDblArray[0], BottomFaceBoxDblArray[1], BottomFaceBoxDblArray[2],
                            dBotFaceNormalVecData[0], dBotFaceNormalVecData[1], dBotFaceNormalVecData[2],
                            TopFaceBoxDblArray, ref iGetFarmostPt, ref dShortestDistance, ref dRefDistance);

            //double dZOfTopFace = Math.Abs(dShortestDistance) + 0.003;
            double dZOfTopFace = Math.Abs(dShortestDistance) + dTopOffsetValue/1000;

            CreateFacetForRectangleExtrude(dXLowerOfBottomFace, dYLowerOfBottomFace, dZLowerOfBottomFace, dXUpperOfBottomFace, dYUpperOfBottomFace, dZOfTopFace);
            
            #region  Commented code 
            /**
            IMathPoint varProjectedPt = null;
            IMathPoint varTopFacePt1 = null;
            if (iGetFarmostPt == 1)
            {
                varProjectedPt = mathUtil.CreatePoint(new double[] { (TopFaceBoxDblArray[0] - dBotFaceNormalVecData[0] * dRefDistance),
                                                                    (TopFaceBoxDblArray[1] - dBotFaceNormalVecData[1] * dRefDistance),
                                                                    (TopFaceBoxDblArray[2] - dBotFaceNormalVecData[2] * dRefDistance) }) as IMathPoint;
                varTopFacePt1 = mathUtil.CreatePoint(new double[] { (TopFaceBoxDblArray[0]), (TopFaceBoxDblArray[1]), (TopFaceBoxDblArray[2]) }) as IMathPoint;
            }
            else
            {
                varProjectedPt = mathUtil.CreatePoint(new double[] { (TopFaceBoxDblArray[3] - dBotFaceNormalVecData[0] * dRefDistance),
                                                                    (TopFaceBoxDblArray[4] - dBotFaceNormalVecData[1] * dRefDistance),
                                                                    (TopFaceBoxDblArray[5] - dBotFaceNormalVecData[2] * dRefDistance) }) as IMathPoint;
                varTopFacePt1 = mathUtil.CreatePoint(new double[] { (TopFaceBoxDblArray[3]), (TopFaceBoxDblArray[4]), (TopFaceBoxDblArray[5]) }) as IMathPoint;
            }

            IMathPoint objProjectPt = varProjectedPt as IMathPoint;
            object objProjPtData = objProjectPt.ArrayData;
            double[] dProjPtData = new double[4];
            dProjPtData = (double[])objProjPtData;
            Debug.Print(" Project Point coord = " + "(" + dProjPtData[0] * 1000.0 + ", " + dProjPtData[1] * 1000.0 + ", " + dProjPtData[2] * 1000.0 + ") mm");

            IMathPoint objTopFacePt1 = varTopFacePt1 as IMathPoint;
            IMathVector vecFromBotToTopFace = objTopFacePt1.Subtract(objProjectPt);
            if (vecFromBotToTopFace == null)
            {
                setTopFace(null);
                setBottomFace(null);
                return;
            }
            double dBeforeNormalization1 = vecFromBotToTopFace.GetLength();
            vecFromBotToTopFace = vecFromBotToTopFace.Normalise();
            double dAfterNormalization1 = vecFromBotToTopFace.GetLength();
            **/
            #endregion

            ///////////////////////////////////////////////////////////////////////////////////
            ///
            #region  Commented code 
            /**
            ModelDoc2 swModel = (ModelDoc2)iSwApp.ActiveDoc;
            if (swModel == null)
            {
                setTopFace(null);
                setBottomFace(null);
                return;
            }

            swModel.InsertSketch2(false);
            //swModel.SketchRectangle(0, 0, 0, .1, .1, .1, false);
            swModel.SketchRectangle(dXLowerOfBottomFace, dYLowerOfBottomFace, dZLowerOfBottomFace, dXUpperOfBottomFace, dYUpperOfBottomFace, dZUpperOfBottomFace, false);

            // Sleep for two seconds
            ////System.Threading.Thread.Sleep(2000);

            //Once sketch gets created, then get  the normal of sketch because in that direction only feature will get create.
            Sketch swSketch = default(Sketch);
            swSketch = swModel.GetActiveSketch2();
            if (swSketch == null)
            {
                MessageBox.Show("Failed to get pointer to 3DSketch1.");
                setTopFace(null);
                setBottomFace(null);
                return;
            }
            //swSketch.SetWorkingPlaneOrientation();
            int refEntityType = 0;
            bool bIfSameDir = false;
            object objSktPlan = swSketch.GetReferenceEntity(ref refEntityType);
            if (objSktPlan is IFace)
            {
                IFace objFace = objSktPlan as IFace;
                object objBottomFaceNormal2 = null;
                objBottomFaceNormal2 = objFace.Normal;
                double[] dBottomFaceNormalValues2 = new double[4];
                dBottomFaceNormalValues2 = (double[])objBottomFaceNormal2;
                //var dirFaceVec2 = mathUtil.CreateVector(new double[] { dBottomFaceNormalValues2[0], dBottomFaceNormalValues2[1], dBottomFaceNormalValues2[2] }) as IMathVector;
                IMathVector dirFaceVec2 = mathUtil.CreateVector(new double[] { dBottomFaceNormalValues2[0], dBottomFaceNormalValues2[1], dBottomFaceNormalValues2[2] }) as IMathVector;
                if (dirFaceVec2 == null)
                {
                    setTopFace(null);
                    setBottomFace(null);
                    return;
                }
                dirFaceVec2 = dirFaceVec2.Normalise();

                object vecFromBotToTopFaceData = vecFromBotToTopFace.ArrayData;
                double[] dFromBotToTopFaceData = new double[4];
                dFromBotToTopFaceData = (double[])vecFromBotToTopFaceData;
                Debug.Print(" Bottom To Top unit vector = " + "(" + dFromBotToTopFaceData[0] + ", " + dFromBotToTopFaceData[1] + ", " + dFromBotToTopFaceData[2] + ")");

                object dirFaceVec2Data = dirFaceVec2.ArrayData;
                double[] dirFaceVec2DataData = new double[4];
                dirFaceVec2DataData = (double[])dirFaceVec2Data;
                Debug.Print(" Bottom Sketch face normal unit vector = " + "(" + dirFaceVec2DataData[0] + ", " + dirFaceVec2DataData[1] + ", " + dirFaceVec2DataData[2] + ")");

                IMathVector vecCrossProduct = vecFromBotToTopFace.Cross(dirFaceVec2);

                if (MyGeometricCompare.MyEqualCompareFunction.bIsEqual(vecCrossProduct.GetLength(), 0.0))
                {
                    double dDotProduct = vecFromBotToTopFace.Dot(dirFaceVec2);
                    if (dDotProduct > 1.0)
                        dDotProduct = 1.0;
                    else if (dDotProduct < -1.0)
                        dDotProduct = -1.0;

                    if (MyGeometricCompare.MyEqualCompareFunction.bIsEqual(dDotProduct, 1))
                        bIfSameDir = true;
                    else
                        bIfSameDir = false;
                }

            }
            else if (objSktPlan is IRefPlane)
            {
                IRefPlane objRefPlan = objSktPlan as IRefPlane;

                MathTransform transform2 = objRefPlan.Transform;
                //transform.
                var transform = objRefPlan.Transform;
                var transform3 = objRefPlan.GetRefPlaneParams();
                var dirVecRefPlane = mathUtil.CreateVector(new double[] { 0, 0, 1 }) as IMathVector;
                //MathVector vecNormalRefPlane = dirVecRefPlane.MultiplyTransform(transform2);
                dirVecRefPlane = dirVecRefPlane.MultiplyTransform(transform2);
                dirVecRefPlane = dirVecRefPlane.Normalise();

                IMathVector vecCrossProduct = vecFromBotToTopFace.Cross(dirVecRefPlane);

                if (MyGeometricCompare.MyEqualCompareFunction.bIsEqual(vecCrossProduct.GetLength(), 0.0))
                {
                    double dDotProduct = vecFromBotToTopFace.Dot(dirVecRefPlane);
                    if (dDotProduct > 1.0)
                        dDotProduct = 1.0;
                    else if (dDotProduct < -1.0)
                        dDotProduct = -1.0;

                    if (MyGeometricCompare.MyEqualCompareFunction.bIsEqual(dDotProduct, 1))
                        bIfSameDir = true;
                    else
                        bIfSameDir = false;
                }
            }

            return;

            bool bToFlipExtrusion = false;
            if (false == bIfSameDir)
            {
                bToFlipExtrusion = true;
            }
            //Extrude the sketch
            IFeatureManager featMan = swModel.FeatureManager;
            IFeature featObj = featMan.FeatureExtrusion(true,
                false, bToFlipExtrusion,
                (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                dZOfTopFace, 0.0,
                false, false,
                false, false,
                0.0, 0.0,
                false, false,
                false, false,
                false,
                false, false);

            if (featObj == null)
            {
                setTopFace(null);
                setBottomFace(null);
                return;
            }
            //////////////////////////////////////////////////////////////
            // Sleep for two seconds
            System.Threading.Thread.Sleep(2000);

            //bool bIsFirstTime = true;
            int iFaceCount = featObj.GetFaceCount();
            var faces = featObj.GetFaces() as object[];
            if (faces != null)
            {
                foreach (IFace2 faceObj in faces)
                {
                    float[] tessTriangs;
                    float[] tessNorms;
                    int iTrangleCount = faceObj.GetTessTriangleCount();
                    tessNorms = faceObj.GetTessNorms() as float[];
                    tessTriangs = faceObj.GetTessTriangles(true) as float[];

                    m_TessTotalTriangsList.AddRange(tessTriangs);
                    m_TessTotalNormsList.AddRange(tessNorms);

                }

            }  
            **/
            #endregion


        }

        #endregion


    }
}
