using System;
using System.ComponentModel;
using System.Data;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using Microsoft.Office.Server.UserProfiles;
using Microsoft.Office.Server.Administration;
using Microsoft.Office.Server;

namespace EmployeeDirectory2010.EmployeeDirectory2010WebPart
{
    [ToolboxItemAttribute(false)]
    public class EmployeeDirectory2010WebPart : WebPart
    {
        private bool _error = false;
        private int _noOfPages = 0;

        private ObjectDataSource ds;

        private SPGridView oGrid;

        private TextBox txtName;
        private TextBox txtDepartment;
        private TextBox txtTitle;
        private Button btnSearch;
        private Button btnClear;

        [Personalizable(PersonalizationScope.Shared)]
        [WebBrowsable(true)]
        [System.ComponentModel.Category("Employee Directory Settings")]
        [WebDisplayName("No Of Pages")]
        [WebDescription("No Of Pages")]
        public int noOfPages
        {
            get
            {
                if (_noOfPages == 0)
                {
                    _noOfPages = 10;
                }
                return _noOfPages;
            }
            set { _noOfPages = value; }
        }


        public EmployeeDirectory2010WebPart()
        {
            this.ExportMode = WebPartExportMode.All;
        }

        protected override void RenderContents(HtmlTextWriter writer)
        {
            //name textbox 
            writer.Write("<table border='0'  width='600px'>");
            writer.Write("<tr><td align='left' nowrap='nowrap' width='20%'>Name : </td>");
            writer.Write("<td align='left' >");
            txtName.RenderControl(writer);
            writer.Write("</td></tr>");

            //department textbox 
            writer.Write("<tr><td align='left' nowrap='nowrap'>Department : </td>");
            writer.Write("<td align='left' >");
            txtDepartment.RenderControl(writer);
            writer.Write("</td></tr>");

            //Title textbox 
            writer.Write("<tr><td align='left' nowrap='nowrap'>Title : </td>");
            writer.Write("<td align='left' >");
            txtTitle.RenderControl(writer);
            writer.Write("</td></tr>");

            // BUTTONS
            writer.Write("<tr><td>");
            btnSearch.RenderControl(writer);
            writer.Write("</td><td>");
            btnClear.RenderControl(writer);
            writer.Write("</td></tr></table>");

            // GRID
            writer.Write("<table border='0'>");
            writer.Write("<tr><td>");
            oGrid.RenderControl(writer);
            writer.Write("</td></tr>");
            writer.Write("</table>");
        }


        protected sealed override void Render(HtmlTextWriter writer)
        {
            oGrid.DataBind();
            if (oGrid.HeaderRow != null)
                foreach (TableCell cell in oGrid.HeaderRow.Cells)
                    cell.CssClass = "GridViewHeaderStyle";

            base.Render(writer);
        }
        /// <summary>
        /// Create all your controls here for rendering.
        /// Try to avoid using the RenderWebPart() method.
        /// </summary>
        protected override void CreateChildControls()
        {
            if (!_error)
            {
                try
                {
                    const string GRIDID = "grid";
                    const string DATASOURCEID = "gridDS";

                    Microsoft.SharePoint.WebControls.CssLink cssLink = new Microsoft.SharePoint.WebControls.CssLink();
                    cssLink.DefaultUrl = "/_layouts/EmployeeDirectory2010/styles/grid.css";
                    this.Page.Header.Controls.Add(cssLink);

                    btnSearch = new Button();
                    btnClear = new Button();
                    txtName = new TextBox();
                    txtDepartment = new TextBox();
                    txtTitle = new TextBox();

                    this.btnSearch.Text = "Search";
                    this.btnSearch.Click += new EventHandler(btnSearch_Click);
                    this.btnClear.Text = "Clear";
                    this.btnClear.Click += new EventHandler(btnClear_Click);

                    ds = new ObjectDataSource();
                    ds.ID = DATASOURCEID;
                    ds.SelectMethod = "SelectData";
                    ds.TypeName = this.GetType().AssemblyQualifiedName;
                    ds.ObjectCreating += new ObjectDataSourceObjectEventHandler(ds_ObjectCreating);
                    this.Controls.Add(ds);

                    this.oGrid = new SPGridView();
                    this.oGrid.ID = GRIDID;
                    oGrid.DataSourceID = ds.ID;
                    this.oGrid.AutoGenerateColumns = false;
                    //this.oGrid.HeaderRow.BackColor = System.Drawing.Color.Blue;
                    //this.oGrid.HeaderRow.ForeColor = System.Drawing.Color.White;
                    //<HeaderRow CssClass="ms-viewheadertr ms-vhltr"/>
                    //<HeaderStyle CssClass="ms-vh2" />
                    //<RowStyle BackColor="ms-itmhover" />
                    //<AlternatingRowStyle CssClass="ms-alternating ms-itmhover" />
                    //<ControlStyle CssClass="ms-vb-title"/>
                    //this.oGrid.HeaderRow.CssClass = "ms-viewheadertr ms-vhltr";
                    //this.oGrid.HeaderStyle.CssClass = "ms-vh2";
                    //this.oGrid.RowStyle.CssClass = "ms-itmhover";
                    //this.oGrid.AlternatingRowStyle.CssClass = "ms-alternating ms-itmhover";
                    //this.oGrid.ControlStyle.CssClass = "ms-vb-title";

                    #region Sorting

                    this.oGrid.AllowSorting = true;

                    ImageField colPicture = new ImageField();

                    colPicture.DataImageUrlField = "Picture";
                    colPicture.HeaderText = "Picture";
                    colPicture.ControlStyle.Width = Unit.Pixel(50);
                    colPicture.ControlStyle.Height = Unit.Pixel(50);
                    colPicture.Visible = true;

                    this.oGrid.Columns.Add(colPicture);

                    BoundField colName = new BoundField();

                    colName.DataField = "Name";
                    colName.HeaderText = "Name";
                    colName.SortExpression = "Name";
                    colName.Visible = false;

                    this.oGrid.Columns.Add(colName);

                    SPMenuField colMenu = new SPMenuField();
                    colMenu.HeaderText = "Name";
                    colMenu.TextFields = "Name";
                    colMenu.SortExpression = "Name";
                    colMenu.MenuTemplateId = "nameMenu";
                    colMenu.NavigateUrlFields = "Uri";
                    colMenu.NavigateUrlFormat = "{0}";
                    colMenu.TokenNameAndValueFields = "Uri=Uri";

                    MenuTemplate titleMenu = new MenuTemplate();
                    titleMenu.ID = "nameMenu";
                    MenuItemTemplate viewMenu = new MenuItemTemplate("View Profile");
                    viewMenu.ClientOnClickNavigateUrl = "%Uri%";
                    titleMenu.Controls.Add(viewMenu);

                    this.Controls.Add(titleMenu);
                    this.oGrid.Columns.Add(colMenu);

                    BoundField colTitle = new BoundField();

                    colTitle.DataField = "Title";
                    colTitle.HeaderText = "Title";
                    colTitle.SortExpression = "Title";
                    colTitle.Visible = true;

                    this.oGrid.Columns.Add(colTitle);

                    BoundField colDepartment = new BoundField();

                    colDepartment.DataField = "Department";
                    colDepartment.HeaderText = "Department";
                    colDepartment.SortExpression = "Department";
                    colDepartment.Visible = true;

                    this.oGrid.Columns.Add(colDepartment);
                    

                    BoundField colEmail = new BoundField();

                    colEmail.DataField = "Email";
                    colEmail.HeaderText = "Email";
                    colEmail.SortExpression = "Email";
                    colEmail.Visible = true;

                    this.oGrid.Columns.Add(colEmail);

                    BoundField colWorkPhone = new BoundField();

                    colWorkPhone.DataField = "WorkPhone";
                    colWorkPhone.HeaderText = "Work Phone";
                    colWorkPhone.SortExpression = "WorkPhone";
                    colWorkPhone.Visible = true;

                    this.oGrid.Columns.Add(colWorkPhone);

                    BoundField colCellPhone = new BoundField();

                    colCellPhone.DataField = "CellPhone";
                    colCellPhone.HeaderText = "Cell Phone";
                    colCellPhone.SortExpression = "CellPhone";
                    colCellPhone.Visible = true;

                    this.oGrid.Columns.Add(colCellPhone);

                    #endregion Sorting

                    #region Paging

                    oGrid.PageSize = noOfPages;
                    oGrid.AllowPaging = true;
                    oGrid.PageIndexChanging += new GridViewPageEventHandler(oGrid_PageIndexChanging);
                    oGrid.PagerTemplate = null;  // Must be called after Controls.Add(oGrid)        

                    #endregion Paging

                    #region filerting
                    // Filtering
                    oGrid.AllowFiltering = true;
                    oGrid.FilterDataFields = ",,,Title,Department";
                    oGrid.FilteredDataSourcePropertyName = "FilterExpression";
                    oGrid.FilteredDataSourcePropertyFormat = "{1} = '{0}'";
                    oGrid.RowDataBound += new GridViewRowEventHandler(oGrid_RowDataBound);
                    ds.Filtering += new ObjectDataSourceFilteringEventHandler(ds_Filtering);
                    oGrid.Sorting += new GridViewSortEventHandler(oGrid_Sorting);

                    #endregion filtering

                    this.Controls.Add(txtName);
                    this.Controls.Add(txtDepartment);
                    this.Controls.Add(txtTitle);
                    this.Controls.Add(btnSearch);
                    this.Controls.Add(btnClear);

                    this.Controls.Add(this.oGrid);

                    oGrid.PagerTemplate = null;

                    oGrid.DataBind();

                    //this.oGrid.HeaderRow.CssClass = "ms-viewheadertr ms-vhltr";
                    this.oGrid.HeaderStyle.CssClass = "GridViewHeaderStyle";
                    this.oGrid.RowStyle.CssClass = "GridRowStyle";
                    this.oGrid.AlternatingRowStyle.CssClass = "GridViewAltRowStyle";
                    //this.oGrid.ControlStyle.CssClass = "ms-vb-title";

                    base.CreateChildControls();
                }
                catch (Exception ex)
                {
                    HandleException(ex);
                }
            }
        }

        void btnClear_Click(object sender, EventArgs e)
        {
            txtName.Text = string.Empty;
            txtDepartment.Text = string.Empty;
            txtTitle.Text = string.Empty;
            oGrid.DataBind();
            //ds.FilterExpression = "Name LIKE '%" + txtName.Text + "%' AND Department LIKE '%" + txtDepartment.Text + "%' AND Title LIKE '%" + txtTitle.Text + "%'";
        }

        void btnSearch_Click(object sender, EventArgs e)
        {
            ds.FilterExpression = "Name LIKE '%" + txtName.Text + "%' AND Department LIKE '%" + txtDepartment.Text + "%' AND Title LIKE '%" + txtTitle.Text + "%'";
        }

        void oGrid_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (sender == null || e.Row.RowType != DataControlRowType.Header)
            {
                return;
            }

            SPGridView grid = sender as SPGridView;

            if (String.IsNullOrEmpty(grid.FilterFieldName))
            {
                return;
            }

            // Show icon on filtered column
            for (int i = 0; i < grid.Columns.Count; i++)
            {
                DataControlField field = grid.Columns[i];

                if (field.SortExpression == grid.FilterFieldName)
                {
                    Image filterIcon = new Image();
                    filterIcon.ImageUrl = "/_layouts/images/filter.gif";
                    filterIcon.Style[HtmlTextWriterStyle.MarginLeft] = "2px";

                    Literal headerText = new Literal();
                    headerText.Text = field.HeaderText;

                    PlaceHolder panel = new PlaceHolder();
                    panel.Controls.Add(headerText);
                    panel.Controls.Add(filterIcon);

                    e.Row.Cells[i].Controls[0].Controls.Add(panel);

                    break;
                }
            }

        }

        void ds_Filtering(object sender, ObjectDataSourceFilteringEventArgs e)
        {
            ViewState["FilterExpression"] = ((ObjectDataSourceView)sender).FilterExpression;
        }

        void ds_ObjectCreating(object sender, ObjectDataSourceEventArgs e)
        {
            e.ObjectInstance = this;
        }


        void oGrid_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            oGrid.PageIndex = e.NewPageIndex;
            oGrid.DataBind();
        }

        void oGrid_Sorting(object sender, GridViewSortEventArgs e)
        {
            if (ViewState["FilterExpression"] != null)
            {
                ds.FilterExpression = (string)ViewState["FilterExpression"];
            }
        }

        protected sealed override void LoadViewState(object savedState)
        {
            base.LoadViewState(savedState);

            if (Context.Request.Form["__EVENTARGUMENT"] != null &&
                 Context.Request.Form["__EVENTARGUMENT"].EndsWith("__ClearFilter__"))
            {
                // Clear FilterExpression
                ViewState.Remove("FilterExpression");
            }
        }

        /// <summary>
        /// Clear all child controls and add an error message for display.
        /// </summary>
        /// <param name="ex"></param>
        private void HandleException(Exception ex)
        {
            this._error = true;
            this.Controls.Clear();
            this.Controls.Add(new LiteralControl(ex.Message));
        }

        public DataTable SelectData()
        {
            try
            {
                DataTable dt = new DataTable("Employees");
                dt.Columns.Add("Picture");
                dt.Columns.Add("Name");
                dt.Columns.Add("Uri");
                dt.Columns.Add("Title");
                dt.Columns.Add("Department");
                dt.Columns.Add("Email");
                dt.Columns.Add("WorkPhone");
                dt.Columns.Add("CellPhone");

                DataRow dr;

                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    using (SPSite site = new SPSite(SPContext.Current.Site.ID))
                    {
                        SPServiceContext serviceContext = SPServiceContext.GetContext(site);

                        //initialize user profile config manager object
                        UserProfileManager profileManager = new UserProfileManager(serviceContext);
                        // UserProfileManager profileManager = new UserProfileManager(ServerContext.Current);

                        List<Employees> employees = new List<Employees>();

                        foreach (UserProfile profile in profileManager)
                        {

                            if (profile != null)
                            {
                                Employees emp = new Employees();
                                emp.Name = (String)profile[PropertyConstants.PreferredName].Value;
                                emp.LastName = (String)profile[PropertyConstants.LastName].Value;
                                if (profile.PublicUrl !=null)
                                {
                                    emp.MySiteUri = (String)profile.PublicUrl.AbsoluteUri;
                                }
                                emp.FirstName = (String)profile[PropertyConstants.FirstName].Value;
                                emp.PictureURL = (String)profile[PropertyConstants.PictureUrl].Value;
                                emp.Title = (String)profile[PropertyConstants.Title].Value;
                                emp.Department = (String)profile[PropertyConstants.Department].Value;
                                emp.WorkEmail = (String)profile[PropertyConstants.WorkEmail].Value;
                                emp.WorkPhone = (String)profile[PropertyConstants.WorkPhone].Value;
                                emp.CellPhone = (String)profile[PropertyConstants.CellPhone].Value;
                                employees.Add(emp);
                                //employees.Add(new Employees()
                                //{
                                //    Name = (String)profile[PropertyConstants.PreferredName].Value,
                                //    LastName = (String)profile[PropertyConstants.LastName].Value,
                                //    MySiteUri = (String)profile.PublicUrl.AbsoluteUri,
                                //    FirstName = (String)profile[PropertyConstants.FirstName].Value,
                                //    PictureURL = (String)profile[PropertyConstants.PictureUrl].Value,
                                //    Title = (String)profile[PropertyConstants.Title].Value,
                                //    Department = (String)profile[PropertyConstants.Department].Value,
                                //    WorkEmail = (String)profile[PropertyConstants.WorkEmail].Value,
                                //    WorkPhone = (String)profile[PropertyConstants.WorkPhone].Value,
                                //    CellPhone = (String)profile[PropertyConstants.CellPhone].Value
                                //});
                            }

                        }

                        foreach (Employees employee in employees)
                        {
                            dr = dt.NewRow();
                            if (string.IsNullOrEmpty(employee.PictureURL))
                            {
                                dr["Picture"] = "~/_layouts/Images/EmployeeDirectory2010/no_image.gif";
                            }
                            else
                            {
                                dr["Picture"] = employee.PictureURL;
                            }
                            dr["Name"] = employee.Name;
                            dr["Uri"] = employee.MySiteUri;
                            dr["Title"] = employee.Title;
                            dr["Department"] = employee.Department;
                            dr["Email"] = employee.WorkEmail;
                            dr["WorkPhone"] = employee.WorkPhone;
                            dr["CellPhone"] = employee.CellPhone;
                            dt.Rows.Add(dr);
                        }
                    }
                });

                DataView v = dt.DefaultView;
                v.Sort = "Name ASC";
                dt = v.ToTable();
                return dt;
            }
            catch (Exception ex)
            {
                return null;
            }
        }        
    }
    

}