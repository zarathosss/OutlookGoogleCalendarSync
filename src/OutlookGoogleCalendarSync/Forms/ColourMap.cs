using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class ColourMap : Form {

        private static readonly ILog log = LogManager.GetLogger(typeof(ColourMap));
        public static Extensions.OutlookColourPicker OutlookComboBox = new Extensions.OutlookColourPicker();
        public static Extensions.GoogleColourPicker GoogleComboBox = new Extensions.GoogleColourPicker();
        
        public ColourMap() {
            OutlookComboBox = null;
            OutlookComboBox = new Extensions.OutlookColourPicker();
            OutlookComboBox.AddCategoryColours();
            GoogleComboBox = null;
            GoogleComboBox = new Extensions.GoogleColourPicker();
            GoogleComboBox.AddPaletteColours();

            InitializeComponent();
            initialiseDataGridView();
        }

        private void initialiseDataGridView() {
            try {
                log.Info("Opening colour mapping window.");
                loadConfig();
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }
        
        private void loadConfig() {
            try {
                if (Settings.Instance.ColourMaps.Count > 0) colourGridView.Rows.Clear();
                foreach (KeyValuePair<String, String> colourMap in Settings.Instance.ColourMaps) {
                    addRow(colourMap.Key, GoogleOgcs.EventColour.Palette.GetColourName(colourMap.Value));
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Populating gridview cells from Settings.", ex);
            }
        }

        private void addRow(String outlookColour, String googleColour) {
            int lastRow = 0;
            try {
                lastRow = colourGridView.Rows.GetLastRow(DataGridViewElementStates.None);
                Object currentValue = colourGridView.Rows[lastRow].Cells["OutlookColour"].Value;
                if (currentValue != null && currentValue.ToString() != "") {
                    lastRow++;
                    colourGridView.Rows.Insert(lastRow);
                }
                colourGridView.Rows[lastRow].Cells["OutlookColour"].Value = outlookColour;
                colourGridView.Rows[lastRow].Cells["GoogleColour"].Value = googleColour;

                colourGridView.CurrentCell = colourGridView.Rows[lastRow].Cells[1];
                colourGridView.NotifyCurrentCellDirty(true);
                colourGridView.NotifyCurrentCellDirty(false);

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Adding colour/category map row #" + lastRow, ex);
            }
        }
        
        #region EVENTS
        private void btSave_Click(object sender, EventArgs e) {
            try {
                Settings.Instance.ColourMaps.Clear();
                foreach (DataGridViewRow row in colourGridView.Rows) {
                    if (row.Cells[0].Value == null || row.Cells[0].Value.ToString().Trim() == "") continue;
                    try {
                        Settings.Instance.ColourMaps.Add(row.Cells[0].Value.ToString(), GoogleOgcs.EventColour.Palette.GetColourId(row.Cells[1].Value.ToString()));                        
                    } catch (System.ArgumentException ex) {
                        if (OGCSexception.GetErrorCode(ex) == "0x80070057") {
                            //An item with the same key has already been added
                        } else throw;
                    }
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Could not save colour/category mappings to Settings.", ex);
            } finally {
                this.Close();
            }
        }

        private void colourGridView_DataError(object sender, DataGridViewDataErrorEventArgs e) {
            /*log.Error(e.Context.ToString());
            if (e.Exception.HResult == -2147024809) { //DataGridViewComboBoxCell value is not valid.
                DataGridViewCell cell = colourGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
                log.Warn("Cell[" + cell.RowIndex + "][" + cell.ColumnIndex + "] has invalid value of '" + cell.Value + "'. Removing.");
                cell.OwningRow.Cells[0].Value = null;
                cell.OwningRow.Cells[1].Value = null;
            } else {
                try {
                    DataGridViewCell cell = colourGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    log.Debug("Cell[" + cell.RowIndex + "][" + cell.ColumnIndex + "] caused error.");
                } catch {
                } finally {
                    OGCSexception.Analyse("Bad cell value in timezone data grid.", e.Exception);
                }
            }*/
        }

        private void colourGridView_CellClick(object sender, DataGridViewCellEventArgs e) {
            if (!this.Visible) return;

            Boolean validClick = (e.RowIndex != -1 && e.ColumnIndex != -1); //Make sure the clicked row/column is valid.
            //Check to make sure the cell clicked is the cell containing the combobox 
            if (validClick && colourGridView.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn) {
                colourGridView.BeginEdit(true);
                ((ComboBox)colourGridView.EditingControl).DroppedDown = true;
            }
        }
        
        private void colourGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e) {
            try {
                if (e.Control is ComboBox) {
                    ComboBox cb = e.Control as ComboBox;
                    cb.DrawMode = DrawMode.OwnerDrawFixed;
                    cb.SelectedIndexChanged -= colourGridView_SelectedIndexChanged;
                    cb.SelectedIndexChanged += colourGridView_SelectedIndexChanged;
                    if (cb is Extensions.OutlookColourCombobox) {
                        cb.DrawItem -= OutlookComboBox.ColourPicker_DrawItem;
                        cb.DrawItem += OutlookComboBox.ColourPicker_DrawItem;
                        OutlookComboBox.ColourPicker_DrawItem(sender, null);
                    } else if (cb is Extensions.GoogleColourCombobox) {
                        cb.DrawItem -= GoogleComboBox.ColourPicker_DrawItem;
                        cb.DrawItem += GoogleComboBox.ColourPicker_DrawItem;
                        GoogleComboBox.ColourPicker_DrawItem(sender, null);
                    }
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }
        #endregion


        private void colourGridView_SelectedIndexChanged(object sender, EventArgs e) {
            //((ComboBox)sender).BackColor = System.Drawing.Color.Red; // (System.Drawing.Color)((ComboBox)sender).SelectedItem;
            //colourGridView.CurrentCell.ba
        }

        private void colourGridView_CurrentCellDirtyStateChanged(object sender, EventArgs e) {
            //log.Debug("colourGridView_CurrentCellDirtyStateChanged");
            //colourGridView.CurrentCell.Style.BackColor = System.Drawing.Color.Blue;
            DataGridViewColumn col = colourGridView.Columns[colourGridView.CurrentCell.ColumnIndex];
            //if (col is DataGridViewComboBoxColumn) {
            //    colourGridView.CommitEdit(DataGridViewDataErrorContexts.Commit);
            //    colourGridView.EndEdit();
            //}
            //colourGridView.celEditingControl
            
        }

        private void colourGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e) {
            newRowNeeded();
        }

        private void newRowNeeded() {
            int lastRow = 0;
            try {
                lastRow = colourGridView.Rows.GetLastRow(DataGridViewElementStates.None);
                Object currentOValue = colourGridView.Rows[lastRow].Cells["OutlookColour"].Value;
                Object currentGValue = colourGridView.Rows[lastRow].Cells["GoogleColour"].Value;
                if (currentOValue != null && currentOValue.ToString() != "" &&
                    currentGValue != null && currentGValue.ToString() != "")
                {
                    lastRow++;
                    DataGridViewCell lastCell = colourGridView.Rows[lastRow - 1].Cells[1];
                    if (lastCell != colourGridView.CurrentCell)
                        colourGridView.CurrentCell = lastCell;
                    colourGridView.NotifyCurrentCellDirty(true);
                    colourGridView.NotifyCurrentCellDirty(false);
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Adding colour/category map row #" + lastRow, ex);
            }            
        }

        private void colourGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            //log.Debug("CellFormatting");
            
        }

        private void colourGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e) {
            //log.Debug("colourGridView_CellValueChanged");
        }

        private void colourGridView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e) {
            //log.Debug("colourGridView_CellPainting "+ e.RowIndex +":"+ e.ColumnIndex);
            //e.PaintBackground(e.ClipBounds, true);
            //e.PaintContent(e.ClipBounds);
            //e.CellStyle.BackColor = System.Drawing.Color.Red;
            //e.Handled = true;
        }
        
        private void colourGridView_CellEnter(object sender, DataGridViewCellEventArgs e) {
            if (colourGridView.CurrentRow.Index + 1 < colourGridView.Rows.Count) return;

            newRowNeeded();
        }
    }
}
