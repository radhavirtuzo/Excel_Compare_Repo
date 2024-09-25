<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Excel_File_compare.aspx.cs" Inherits="Excel_Compare_File.Excel_File_compare" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Excel Compare</title>
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet" />
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
    <style>

                body {
                    backdrop-filter: blur(10px);
                    padding: 0;
                    margin: 0;
                    box-sizing: border-box;
                }

                .header {
                    background-image: url('img/bread_crumb_bg.png');
                     
                    background-size: cover;
                    background-position: center;
                    position: sticky;
                    top: 0;
                    width: 100%;
                    height: 250px;
                    
                    display: flex;
                    justify-content: center;
                    align-items: baseline;
                    z-index: 999;
                    padding-top: 50px;
                }

                .header::after {
                    content: "";
                    background-image: url('img/inner_page_banner_overlay.svg');
                    position: absolute;
                    bottom: -1px;
                    left: 0;
                    background-size: cover;
                    background-repeat: no-repeat;
                    width: 100%;
                    height: 200px;
                    background-position: center;
                }

                .header h1 {
                    color: #ffffff;
                    text-align: center;
                    
                }

                

                .card {
                    background-color: rgba(255, 255, 255, 0.95);
                    border: transparent;
                    border-radius: 25px;
                    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
                  
                }

                .card:hover {
                    border-color: #007bff;
                    transition: border-color 0.3s;
                }

                .btn-primary, .btn-secondary {
                    transition: background-color 0.3s, transform 0.2s;
                }

                .btn-primary:hover {
                    background-color: #0056b3;
                    transform: scale(1.05);
                }

                .btn-secondary:hover {
                    background-color: #6c757d;
                    transform: scale(1.05);
                }

                .table-container {
                    max-height: 400px;
                    overflow-y: auto;
                }

                .loader {
                    position: fixed;
                    left: 50%;
                    top: 50%;
                    transform: translate(-50%, -50%);
                    z-index: 1000;
                    text-align: center;
                }

                .loader .spinner {
                    width: 50px;
                    height: 50px;
                    border: 5px solid rgba(0, 123, 255, 0.2);
                    border-top-color: rgba(0, 123, 255, 1);
                    border-radius: 50%;
                    animation: spin 1s linear infinite;
                }

                section.excel_folder {
                    min-height: 100%;
                    background-color: #f6f4fe;
                    padding: 20px 0 50px;
                  
                }

                section.excel_folder h5 {
                    font-size: 28px;
                    font-weight: 700;
                  
                }

                section.excel_folder label {
                    color: #dc0808;
                    font-size: 14px;
                    font-weight: 500;
                }

                .loader-text {
                    margin-top: 10px;
                    font-size: 18px;
                    color: #007bff;
                }

                section.excel_folder input {
                    padding: 5px 20px 5px 10px;
                    color: #3E3F66;
                    border: 2px solid #E1DBF4;
                    border-radius: 12px;
                    font-weight: 500;
                    height: auto;
                }

                section.excel_folder .btn {
                    font-size: 14px;
                    padding: 9px 22px;
                    border-radius: 10px;
                    margin-left: 20px;
                    position: relative;
                    border: 1px solid var(--danger);
                    text-align: center;
                }

                section.excel_folder .btn.btn-blue {
                    background: #013954;
                    border-color: #013954;
                    color: white !important;
                }

                section.excel_folder .btn.btn-red {
                    background: #f03f37;
                    border-color: #f03f37;
                    color: white !important;
                }

                section.excel_folder table th {
                    position: sticky;
                    top: 0;
                    background: #33355f;
                    color: #fff;
                }

                section.excel_folder .scroll-body {
                    min-height: auto !important;
                    max-height: 360px !important;
                    height: 250px !important;
                }

                @keyframes spin {
                    0% {
                        transform: rotate(0deg);
                    }

                    100% {
                        transform: rotate(360deg);
                    }
                }
    </style>
            <script>
                $(document).ready(function () {
                    $('#<%= ButtonCompare.ClientID %>').click(function () {
                        $('.loader').show();
                    });
                });
            </script>
</head>
<body>
    <form id="form1" runat="server">
        <div class="header">
            <h1>Excel File Comparison </h1>
        </div>
        <div class="loader" style="display: none;">
            <div class="spinner"></div>
            <div class="loader-text">Processing...</div>
        </div>
        <section class="excel_folder">
            <div class="container">
                <div class="card">
                    <div class="card-body">
                        <h5 style="text-align: center;"><span style="color: #dc3545">Upload </span><span style="color: #32236F;">Files</span></h5>
                        <div class="form-group">
                            <label for="FileUpload1">Upload First Excel File:</label>
                            <asp:FileUpload ID="FileUpload1" runat="server" Accept=".xls,.xlsx" CssClass="form-control" />
                        </div>
                        <div class="form-group">
                            <label for="FileUpload2">Upload Second Excel File:</label>
                            <asp:FileUpload ID="FileUpload2" runat="server" Accept=".xls,.xlsx" CssClass="form-control" />
                        </div>
                        <div class=" text-center">
                            <asp:Button ID="ButtonCompare" runat="server" Text="Compare" CssClass="btn btn-blue" OnClick="ButtonCompare_Click"  />
                            <asp:Button ID="ButtonRefresh" runat="server" Text="Refresh" CssClass="btn btn-red ml-2" OnClick="ButtonRefresh_Click" />
                        </div>
                          <div class="col-md-12 text-center mt-4">
                               <asp:Label ID="LabelResult" runat="server" ForeColor="Red" CssClass="mt-2"></asp:Label>
                        </div>
                    </div>
                </div>

                <div class="card mt-5" runat="server" id="cardresult">

                    <div class="card-body">
                        <div class="row">
                            <div class="col-md-6">
                                <h5><span style="color: #32236F">Comparison </span><span style="color: #dc3545;">Results</span></h5>

                            </div>
                           
                            <div class="col-md-6 " style="text-align: end;">
                                <asp:Button ID="ButtonExport" CssClass="btn btn-blue" runat="server" Text="Export to Excel" OnClick="ButtonExport_Click" />

                            </div>
                        </div>

                        <div class="mt-4 table-responsive">
                            <div class="scroll-body">
                                <asp:GridView ID="GridViewResult" runat="server" AutoGenerateColumns="true" CssClass="table table-striped " Visible="false"  OnRowDataBound="GridViewResult_RowDataBound">
                                    
                                </asp:GridView>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </section>
    </form> 
</body>
</html>
