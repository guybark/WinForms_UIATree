using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Automation.Provider;
using System.Runtime.InteropServices;
using System.Windows.Automation;

namespace WinForms_UIATree
{
    public partial class GardenControl : UserControl,
        IRawElementProviderFragmentRoot,
        IGridProvider,
        ITableProvider,
        IRawElementProviderHwndOverride
    {
        public GardenGroup[] GardenGroups { get; set; }

        private UIARowHeaders[] uiaRowHeaders;

        private Pen gridLinePen = new Pen(SystemColors.WindowText, 2);

        public GardenControl()
        {
            InitializeComponent();

            // Have this UserControl fill the Panel that contains it.
            this.Dock = DockStyle.Fill;
            this.BorderStyle = BorderStyle.FixedSingle;

            // Add a CheckBox near the top left corner of the UserControl.
            var checkBox = new CheckBox();
            this.Controls.Add(checkBox);
            checkBox.Location = new Point(2, 2);

            // Add to demo data.
            InitializeItems();
        }

        private void InitializeItems()
        {
            this.GardenGroups = new GardenGroup[3];

            GardenGroups[0] = new GardenGroup(this, 0, Name = "Birds");
            GardenGroups[0].GardenThings.Add(new GardenThing(GardenGroups[0], 0, "Towhee", 3));
            GardenGroups[0].GardenThings.Add(new GardenThing(GardenGroups[0], 1, "Steller's Jay", 5));
            GardenGroups[0].GardenThings.Add(new GardenThing(GardenGroups[0], 2, "Hummingbird", 2));

            GardenGroups[1] = new GardenGroup(this, 1, "Mammals");
            GardenGroups[1].GardenThings.Add(new GardenThing(GardenGroups[1], 0, "Squirrel", 10));
            GardenGroups[1].GardenThings.Add(new GardenThing(GardenGroups[1], 1, "Racoon", 1));
            GardenGroups[1].GardenThings.Add(new GardenThing(GardenGroups[1], 2, "Opossum", 0));

            GardenGroups[2] = new GardenGroup(this, 2, "Plants");
            GardenGroups[2].GardenThings.Add(new GardenThing(GardenGroups[2], 0, "Jasmine", 1));
            GardenGroups[2].GardenThings.Add(new GardenThing(GardenGroups[2], 1, "Poppy", 4));
            GardenGroups[2].GardenThings.Add(new GardenThing(GardenGroups[2], 2, "Lavender", 40));
        }

        // Render some grid lines in the controls.
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            int yOffset = 0;

            for (int i = 0; i < this.GardenGroups.Length - 1; ++i)
            {
                yOffset += (this.Height / this.GardenGroups.Length);
                e.Graphics.DrawLine(gridLinePen, 0, yOffset, this.Width, yOffset);
            }

            int xOffset = 0;

            for (int i = 0; i < this.GardenGroups[0].GardenThings.Count - 1; ++i)
            {
                xOffset += (this.Width / this.GardenGroups[0].GardenThings.Count);
                e.Graphics.DrawLine(gridLinePen, xOffset, 0, xOffset, this.Height);
            }
        }

        // WndProc() needs to be overridden in order to expose our support for the UIA provider API.
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case NativeUIA.WM_GETOBJECT:
                {
                    // If the window is being asked for a UIA provider, return ourselves.
                    if (m.LParam == NativeUIA.UiaRootObjectId)
                    {
                        m.Result = NativeUIA.UiaReturnRawElementProvider(this.Handle, m.WParam, m.LParam, this);

                        return;
                    }

                    break;
                }
            }

            base.WndProc(ref m);
        }

        #region IRawElementProviderFragmentRoot

        // Return the GardenThing beneath this point.
        public IRawElementProviderFragment ElementProviderFromPoint(double x, double y)
        {
            for (int rowIdx = 0; rowIdx < this.GardenGroups.Length; rowIdx++)
            {
                // Is the point insde this row?
                if (this.GardenGroups[rowIdx].BoundingRectangle.Contains(x, y))
                {
                    var row = this.GardenGroups[rowIdx];

                    for (int cellIdx = 0; cellIdx < row.GardenThings.Count; cellIdx++)
                    {
                        // Is the point inside this cell?
                        if (row.GardenThings[cellIdx].BoundingRectangle.Contains(x, y))
                        {
                            return row.GardenThings[cellIdx];
                        }
                    }
                }
            }

            return this;
        }

        public IRawElementProviderFragment GetFocus()
        {
            // None of the dynamically-created UIA elements support getting keyboard focus.
            return this;
        }

        #endregion // IRawElementProviderFragmentRoot

        #region IRawElementProviderFragment

        public System.Windows.Rect BoundingRectangle
        {
            get
            {
                var screenLocation = this.PointToScreen(this.Location);

                return new System.Windows.Rect(
                            screenLocation.X, screenLocation.Y,
                            this.Right, this.Bottom);
            }
        }

        public IRawElementProviderFragmentRoot FragmentRoot
        {
            get
            {
                // This UserControl is its own FragmentRoot, and UIA can do what it 
                // needs to given that the UserControl is hwnd-based.
                return null;
            }
        }

        public IRawElementProviderSimple[] GetEmbeddedFragmentRoots()
        {
            return null;
        }

        public int[] GetRuntimeId()
        {
            // Implementations should return null for an element that corresponds to a window handle.
            return null;
        }

        public IRawElementProviderFragment Navigate(NavigateDirection direction)
        {
            if (direction == NavigateDirection.FirstChild)
            {
                if (GardenGroups.Length > 0)
                {
                    return this.GardenGroups[0];
                }
            }
            else if (direction == NavigateDirection.LastChild)
            {
                if (GardenGroups.Length > 0)
                {
                    return this.GardenGroups[this.GardenGroups.Length - 1];
                }
            }

            return null;
        }

        public void SetFocus()
        {
            // This implementation does not support programmatically setting
            // keyboard focus to a row in the grid.
        }

        #endregion // IRawElementProviderFragment

        #region IRawElementProviderSimple 

        public IRawElementProviderSimple HostRawElementProvider
        {
            get
            {
                // We don't need to invoke on a different thread because our ProviderOptions
                // returns UseComThreading.
                return AutomationInteropProvider.HostProviderFromHandle(this.Handle);
            }
        }

        public ProviderOptions ProviderOptions
        {
            get
            {
                return ProviderOptions.ServerSideProvider | ProviderOptions.UseComThreading;
            }
        }

        public object GetPatternProvider(int patternId)
        {
            if ((patternId == GridPatternIdentifiers.Pattern.Id) ||
                (patternId == TablePatternIdentifiers.Pattern.Id))
            {
                return this;
            }

            return null;
        }

        public object GetPropertyValue(int propertyId)
        {
            // By default, the element gets exposed through the Control view of the UIA tree,
            // so we don't need to react to IsControlElementProperty here.

            // For this demo, this element is not keyboard focusable. If it were, then it would
            // need to return true for IsKeyboardFocusableProperty, and either true for 
            // HasKeyboardFocusProperty if it has focus now.

            // IMPORTANT: The Name and LocalizedControlType returned throughout this file 
            // should be localized!

            if (propertyId == AutomationElementIdentifiers.NameProperty.Id)
            {
                return "Back";
            }
            else if (propertyId == AutomationElementIdentifiers.LocalizedControlTypeProperty.Id)
            {
                // If the element's control type was a match for a known control type, 
                // (for example, a Button,) then that specific ControlType would be 
                // returned from GetPropertyValue, and LocalizedControlType would not
                // need to be handled here.
                return "Garden";
            }

            return null;
        }

        #endregion // IRawElementProviderSimple

        #region IGridProvider

        public int RowCount
        {
            get
            {
                return this.GardenGroups.Length;
            }
        }

        public int ColumnCount
        {
            get
            {
                // Assume an item exists and all GardenGroups have the same subitem count.
                return this.GardenGroups[0].GardenThings.Count;
            }
        }

        public IRawElementProviderSimple GetItem(int rowIndex, int columnIndex)
        {
            // Return the cell at a specific row and column position in the grid.
            if (rowIndex < this.GardenGroups.Length)
            {
                var row = this.GardenGroups[rowIndex];

                if (columnIndex < row.GardenThings.Count)
                {
                    return row.GardenThings[columnIndex];
                }
            }

            return null;
        }

        #endregion // IGridProvider

        #region ITableProvider

        public RowOrColumnMajor RowOrColumnMajor
        {
            get
            {
                // The table is primarily navigated row-by-row.
                return RowOrColumnMajor.RowMajor;
            }
        }

        public IRawElementProviderSimple[] GetColumnHeaders()
        {
            // This demo control has no column headers.
            return null;
        }

        public IRawElementProviderSimple[] GetRowHeaders()
        {
            // Have we cached the row header names yet? This assumes the order of
            // the row headers cannot change.
            if (this.uiaRowHeaders == null)
            {
                this.uiaRowHeaders = new UIARowHeaders[this.GardenGroups.Length];

                for (int i = 0; i < this.GardenGroups.Length; i++)
                {
                    // Create a UIARowHeaders specifically in order to enable screen readers 
                    // to announce the row header text when at a cell in the control.
                    this.uiaRowHeaders[i] = new UIARowHeaders(this.GardenGroups[i].Name);
                }
            }

            return uiaRowHeaders;
        }

        #endregion // ITableProvider

        #region IRawElementProviderHwndOverride

        public IRawElementProviderSimple GetOverrideProviderForHwnd(IntPtr hwnd)
        {
            // In this demo app, there's only has one child hwnd off the UserControl, and it's fixed in the 
            // first cell in the UserControl. If the UserControl had multiple child hwnds, we'd need make to 
            // the connection here between child hwnd supplied to GetOverrideProviderForHwnd() and the cell 
            // that contains it.

            // For this demo, return the first cell. Assume we've created some content.
            return this.GardenGroups[0].GardenThings[0];
        }

        #endregion // IRawElementProviderHwndOverride

        #region UIARowHeaders

        // The UIARowHeaders class exists only to provide row header text to the GardenControl
        // and the GardenThing when they're asked to provide the text of row headers.

        // While a UIA client will announce the text returned by the UIARowHeaders, no
        // other commands can be performed through it. For example, Narrator could not
        // navigate between adjacent row headers based on the UIARowHeaders.

        private class UIARowHeaders : IRawElementProviderSimple
        {
            private string rowHeaderText;

            public UIARowHeaders(string rowHeaderText)
            {
                this.rowHeaderText = rowHeaderText;
            }

            public IRawElementProviderSimple HostRawElementProvider
            {
                get
                {
                    return null;
                }
            }

            public ProviderOptions ProviderOptions
            {
                get
                {
                    return ProviderOptions.ServerSideProvider | ProviderOptions.UseComThreading;
                }
            }

            public object GetPatternProvider(int patternId)
            {
                return null;
            }

            public object GetPropertyValue(int propertyId)
            {
                if (propertyId == AutomationElementIdentifiers.NameProperty.Id)
                {
                    return this.rowHeaderText;
                }

                return null;
            }
        }

        #endregion // GridCellColumnHeaderUIA
    }

    public class GardenGroup :
        IRawElementProviderFragment
    {
        public string Name { get; set; }

        private List<GardenThing> gardenThings = new List<GardenThing>();
        public List<GardenThing> GardenThings { get => gardenThings; }

        private GardenControl parent;
        public GardenControl GardenControl { get => this.parent; }

        private int itemIndex;
        public int ItemIndex { get => itemIndex; }

        public GardenGroup(GardenControl parent, int itemIndex, string name)
        {
            this.parent = parent;
            this.itemIndex = itemIndex;
            this.Name = name;
        }

        #region IRawElementProviderFragment

        public System.Windows.Rect BoundingRectangle
        {
            get
            {
                var gardenControlBounds = this.parent.BoundingRectangle;

                var itemHeight = gardenControlBounds.Height / this.parent.GardenGroups.Length;
                var itemYOffset = this.itemIndex * itemHeight;

                return new System.Windows.Rect(
                    gardenControlBounds.Left,
                    gardenControlBounds.Top + itemYOffset,
                    gardenControlBounds.Width,
                    itemHeight);
            }
        }

        public IRawElementProviderFragmentRoot FragmentRoot
        {
            get
            {
                return this.parent;
            }
        }

        public IRawElementProviderSimple[] GetEmbeddedFragmentRoots()
        {
            return null;
        }

        public int[] GetRuntimeId()
        {
            // Give the GardenGroup a unique id amongst its peers, and different from its ancestors.
            return new int[] {
                    AutomationInteropProvider.AppendRuntimeId,
                    this.ItemIndex + 1
                };
        }

        public IRawElementProviderFragment Navigate(NavigateDirection direction)
        {
            // UIA can navigate up to the parent based on the hwnd relationships.

            if (direction == NavigateDirection.Parent)
            {
                return this.parent;
            }
            else if (direction == NavigateDirection.PreviousSibling)
            {
                if (this.itemIndex > 0)
                {
                    return this.parent.GardenGroups[this.itemIndex - 1];
                }
            }
            else if (direction == NavigateDirection.NextSibling)
            {
                if (this.itemIndex < this.parent.GardenGroups.Length - 1)
                {
                    return this.parent.GardenGroups[this.itemIndex + 1];
                }
            }
            else if (direction == NavigateDirection.FirstChild)
            {
                if (this.GardenThings.Count > 0)
                {
                    return this.GardenThings[0];
                }
            }
            else if (direction == NavigateDirection.LastChild)
            {
                if (this.GardenThings.Count > 0)
                {
                    return this.GardenThings[this.GardenThings.Count - 1];
                }
            }

            return null;
        }

        public void SetFocus()
        {
            // This implementation does not support programmatically setting
            // keyboard focus to a row in the grid.
        }

        #endregion // IRawElementProviderFragment

        #region IRawElementProviderSimple 

        public IRawElementProviderSimple HostRawElementProvider
        {
            get
            {
                return null;
            }
        }

        public ProviderOptions ProviderOptions
        {
            get
            {
                return ProviderOptions.ServerSideProvider | ProviderOptions.UseComThreading;
            }
        }

        public object GetPatternProvider(int patternId)
        {
            return null;
        }

        public object GetPropertyValue(int propertyId)
        {
            // By default, the element gets exposed through the Control view of the UIA tree,
            // so we don't need to react to IsControlElementProperty here.

            // For this demo, this element is not keyboard focusable. If it were, then it would
            // need to return true for IsKeyboardFocusableProperty, and either true for 
            // HasKeyboardFocusProperty is it has focus now.

            if (propertyId == AutomationElementIdentifiers.NameProperty.Id)
            {
                return this.Name;
            }
            else if (propertyId == AutomationElementIdentifiers.LocalizedControlTypeProperty.Id)
            {
                return "Garden Group";
            }

            return null;
        }

        #endregion // IRawElementProviderSimple
    }

    public class GardenThing :
        IRawElementProviderFragment,
        IGridItemProvider,
        ITableItemProvider,
        IValueProvider
    {
        public string Name { get; set; }

        private GardenGroup parent;
        private int subItemIndex;

        public GardenThing(GardenGroup parent, int subItemIndex, string name, int count)
        {
            this.parent = parent;
            this.subItemIndex = subItemIndex;
            this.Name = name;
            this.Value = count.ToString();
        }

        #region IRawElementProviderFragment

        public System.Windows.Rect BoundingRectangle
        {
            get
            {
                var gardenGroupBounds = this.parent.BoundingRectangle;

                var subItemWidth = gardenGroupBounds.Width / this.parent.GardenThings.Count;
                var subItemXOffset = this.subItemIndex * subItemWidth;

                return new System.Windows.Rect(
                    gardenGroupBounds.Left + subItemXOffset,
                    gardenGroupBounds.Top,
                    subItemWidth,
                    gardenGroupBounds.Height);
            }
        }

        public IRawElementProviderFragmentRoot FragmentRoot
        {
            get
            {
                return this.parent.GardenControl;
            }
        }

        public IRawElementProviderSimple[] GetEmbeddedFragmentRoots()
        {
            return null;
        }

        public int[] GetRuntimeId()
        {
            // Give the GardenThing a unique id amongst its peers, and different from its ancestors.
            // Note that any GardenThing that's an overrider provider for a Win32 hwnd will get a
            // runtime id based on the Win32 hwnd, not what's returned here.
            return new int[] { AutomationInteropProvider.AppendRuntimeId,
                this.parent.GardenControl.GardenGroups.Length + 
                (this.parent.GardenThings.Count * this.parent.ItemIndex) +
                this.subItemIndex + 1};
        }

        public IRawElementProviderFragment Navigate(NavigateDirection direction)
        {
            // The GardenThings have no child UIA elements.

            if (direction == NavigateDirection.Parent)
            {
                return this.parent;
            }
            else if (direction == NavigateDirection.PreviousSibling)
            {
                if (this.subItemIndex > 0)
                {
                    return this.parent.GardenThings[this.subItemIndex - 1];
                }
            }
            else if (direction == NavigateDirection.NextSibling)
            {
                if (this.subItemIndex < this.parent.GardenThings.Count - 1)
                {
                    return this.parent.GardenThings[this.subItemIndex + 1];
                }
            }

            return null;
        }

        public void SetFocus()
        {
            // This implementation does not support programmatically setting
            // keyboard focus to a row in the grid.
        }

        #endregion // IRawElementProviderFragment

        #region IRawElementProviderSimple 

        public IRawElementProviderSimple HostRawElementProvider
        {
            get
            {
                // For this demo, the first cell in the control hosts a Win32 hwnd.
                if ((this.Row == 0) && (this.Column == 0))
                {
                    IRawElementProviderSimple result;

                    var hwndCheckBox = this.parent.GardenControl.Controls[0].Handle;

                    NativeUIA.UiaHostProviderFromHwnd(hwndCheckBox, out result);

                    return result;
                }

                return null;
            }
        }

        public ProviderOptions ProviderOptions
        {
            get
            {
                var options = ProviderOptions.ServerSideProvider | ProviderOptions.UseComThreading;

                // For this demo, the first cell in the control hosts a Win32 hwnd.
                if ((this.Row == 0) && (this.Column == 0))
                {
                    options |= ProviderOptions.OverrideProvider;
                }

                return options;
            }
        }

        public object GetPatternProvider(int patternId)
        {
            if ((patternId == GridItemPatternIdentifiers.Pattern.Id) ||
                (patternId == TableItemPatternIdentifiers.Pattern.Id) ||
                (patternId == ValuePatternIdentifiers.Pattern.Id))
            {
                return this;
            }

            if (patternId == ValuePatternIdentifiers.Pattern.Id)
            {
                return this;
            }

            return null;
        }

        public object GetPropertyValue(int propertyId)
        {
            // By default, the element gets exposed through the Control view of the UIA tree,
            // so we don't need to react to IsControlElementProperty here.

            // For this demo, this element is not keyboard focusable. If it were, then it would
            // need to return true for IsKeyboardFocusableProperty, and either true for 
            // HasKeyboardFocusProperty is it has focus now.

            if (propertyId == AutomationElementIdentifiers.NameProperty.Id)
            {
                return this.Name;
            }
            else if (propertyId == AutomationElementIdentifiers.LocalizedControlTypeProperty.Id)
            {
                return "Garden Thing";
            }

            return null;
        }

        #endregion // IRawElementProviderSimple

        #region IGridItemProvider

        public int Row
        {
            get
            {
                // The row index is zero-based.
                return this.parent.ItemIndex;
            }
        }

        public int RowSpan
        {
            get
            {
                // All cells span one row.
                return 1;
            }
        }

        public int Column
        {
            get
            {
                return this.subItemIndex;
            }
        }

        public int ColumnSpan
        {
            get
            {
                // All cells span one column.
                return 1;
            }
        }

        public IRawElementProviderSimple ContainingGrid
        {
            get
            {
                return this.parent.GardenControl;
            }
        }

        #endregion // IGridItemProvider

        #region ITableItemProvider

        public IRawElementProviderSimple[] GetColumnHeaderItems()
        {
            // This demo control has no column headers.
            return null;
        }

        public IRawElementProviderSimple[] GetRowHeaderItems()
        {
            // Return the one row header associated with this cell.
            var uiaRowHeaders = this.parent.GardenControl.GetRowHeaders();
            if ((uiaRowHeaders != null) && (this.subItemIndex < uiaRowHeaders.Length))
            {
                var headers = new IRawElementProviderSimple[1];
                headers[0] = uiaRowHeaders[this.parent.ItemIndex];
                return headers;
            }

            return null;
        }

        #endregion // ITableItemProvider

        #region IValueProvider

        public string Value { get; }

        public bool IsReadOnly { get => true; }

        public void SetValue(string value)
        {
            // Any attempt to set teh value has no effect.
        }

        #endregion // IValueProvider
    }

    // The NativeUIA class provides access to the native UIA provider functions and data.
    public class NativeUIA
    {
        [DllImport("UIAutomationCore.dll", EntryPoint = "UiaReturnRawElementProvider", CharSet = CharSet.Unicode)]
        public static extern IntPtr UiaReturnRawElementProvider(
            IntPtr hwnd, IntPtr wParam, IntPtr lParam, IRawElementProviderSimple el);

        [DllImport("UIAutomationCore.dll", EntryPoint = "UiaHostProviderFromHwnd", CharSet = CharSet.Unicode)]
        public static extern int UiaHostProviderFromHwnd(
            IntPtr hwnd,
            [MarshalAs(UnmanagedType.Interface)] out IRawElementProviderSimple provider);

        public const int WM_GETOBJECT = 0x003D;
        public static IntPtr UiaRootObjectId = (IntPtr)(-25);
    }
}
