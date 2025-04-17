import streamlit as st
import re
import pandas as pd
import os
import json
from io import StringIO, BytesIO
from dotenv import load_dotenv
from openai import AzureOpenAI
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Set page configuration
st.set_page_config(
    page_title="SQL Stored Procedure Analyzer",
    page_icon="🧰",
    layout="wide"
)

# Function to create a Word document from analysis
def create_word_document(analysis):
    # Create a new Document
    doc = Document()

    # Add title
    title = doc.add_heading('SQL Stored Procedure Analysis Report', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Add procedure name
    proc_name_heading = doc.add_heading('Stored Procedure Name:', level=1)
    # Make the procedure name itself bold in its own paragraph
    proc_name_para = doc.add_paragraph()
    proc_name_para.add_run(analysis['procedure_name']).bold = True

    # Add scope
    doc.add_heading('Scope:', level=1)
    doc.add_paragraph(analysis['scope'])

    # Add optimization steps
    doc.add_heading('Optimization Steps:', level=1)

    for i, opt in enumerate(analysis["optimizations"], 1):
        # Step heading
        step_heading = doc.add_heading(f'Step {i}: {opt["type"]}', level=2)

        # Existing Logic
        doc.add_heading('Existing Logic:', level=3)
        existing_code = doc.add_paragraph(opt["existing_logic"])
        # existing_code.style = 'Code'  # <--- REMOVE THIS LINE

        # Format code paragraph
        existing_code_fmt = existing_code.paragraph_format
        existing_code_fmt.left_indent = Inches(0.25)
        existing_code_fmt.right_indent = Inches(0.25)
        # Apply monospaced font to the runs within the paragraph
        for run in existing_code.runs:
            run.font.name = 'Courier New'
            # Optional: Set font size for code
            run.font.size = Pt(10)

        # Optimized Logic
        doc.add_heading('Optimized Logic:', level=3)
        optimized_code = doc.add_paragraph(opt["optimized_logic"])
        # optimized_code.style = 'Code' # <--- REMOVE THIS LINE

        # Format code paragraph
        optimized_code_fmt = optimized_code.paragraph_format
        optimized_code_fmt.left_indent = Inches(0.25)
        optimized_code_fmt.right_indent = Inches(0.25)
        # Apply monospaced font to the runs within the paragraph
        for run in optimized_code.runs:
            run.font.name = 'Courier New'
            # Optional: Set font size for code
            run.font.size = Pt(10)

        # Explanation
        explanation_para = doc.add_paragraph()
        explanation_text = explanation_para.add_run(opt["explanation"])
        explanation_text.italic = True

        # Add separator paragraph (using a paragraph border might be better visually)
        # Consider adding a paragraph border instead of underscores for a cleaner look
        separator = doc.add_paragraph()
        # Example of adding a bottom border (uncomment if desired)
        # p_border = OxmlElement('w:pBdr')
        # bottom_border = OxmlElement('w:bottom')
        # bottom_border.set(qn('w:val'), 'single')
        # bottom_border.set(qn('w:sz'), '6') # size in 1/8ths of a point
        # bottom_border.set(qn('w:space'), '1')
        # bottom_border.set(qn('w:color'), 'auto')
        # p_border.append(bottom_border)
        # separator._p.get_or_add_pPr().append(p_border)
        # Or keep the simple underscore line:
        separator.add_run('_' * 40)


    # Add summary table
    doc.add_heading('Summary:', level=1)

    # Create table data (ensure line_number is handled if missing)
    table_data = []
    for opt in analysis["optimizations"]:
        table_data.append({
            "Type of Change": opt.get("type", "N/A"),
            "Line Number": opt.get("line_number", "N/A"), # Use .get for safety
            "Original Code Snippet": opt.get("existing_logic", "")[:40] + "..." if len(opt.get("existing_logic", "")) > 40 else opt.get("existing_logic", ""),
            "Optimized Code Snippet": opt.get("optimized_logic", "")[:40] + "..." if len(opt.get("optimized_logic", "")) > 40 else opt.get("optimized_logic", ""),
            "Optimization Explanation": opt.get("explanation", "")
        })

    # Add table to document only if there's data
    if table_data:
        num_rows = 1 + len(table_data)
        num_cols = 5
        table = doc.add_table(rows=num_rows, cols=num_cols)
        table.style = 'Table Grid' # Ensure 'Table Grid' style exists or use a known default

        # Set header row
        header_cells = table.rows[0].cells
        headers = ['Type of Change', 'Line Number', 'Original Code Snippet', 'Optimized Code Snippet', 'Optimization Explanation']
        for i, header_text in enumerate(headers):
            cell = header_cells[i]
            # Clear existing content (sometimes needed)
            cell.text = ''
            # Add text and format
            p = cell.paragraphs[0]
            run = p.add_run(header_text)
            run.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER # Center align header text

        # Add data rows
        for i, item in enumerate(table_data):
            row_cells = table.rows[i+1].cells
            row_cells[0].text = item['Type of Change']
            row_cells[1].text = item['Line Number']
            row_cells[2].text = item['Original Code Snippet']
            row_cells[3].text = item['Optimized Code Snippet']
            row_cells[4].text = item['Optimization Explanation']

        # Set table column widths (apply carefully)
        try:
            table.columns[0].width = Inches(1.2)
            table.columns[1].width = Inches(0.8) # Reduced width for line number
            table.columns[2].width = Inches(1.5)
            table.columns[3].width = Inches(1.5)
            table.columns[4].width = Inches(2.0) # Increased width for explanation
        except IndexError:
             st.warning("Could not set all table column widths.")


        # Apply alternate row shading
        for i, row in enumerate(table.rows):
            if i > 0 and i % 2 == 0:  # Even data rows (index 2, 4, ...) get shading
                for cell in row.cells:
                    tcPr = cell._tc.get_or_add_tcPr()
                    shading_elm = OxmlElement('w:shd')
                    shading_elm.set(qn('w:fill'), "F2F2F2") # Light gray color
                    shading_elm.set(qn('w:val'), 'clear') # Ensure fill type is set
                    shading_elm.set(qn('w:color'), 'auto')
                    tcPr.append(shading_elm)
    else:
        doc.add_paragraph("No optimization suggestions were generated.")


    # Save the document to a BytesIO object
    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)

    return doc_io

# --- Keep the rest of your Streamlit code as is ---

# Function to analyze stored procedure using Azure OpenAI
def analyze_stored_procedure(file_content):
    try:
        # Load credentials securely from Streamlit secrets
        azure_openai_endpoint = st.secrets["AZURE_OPENAI_ENDPOINT"]
        azure_openai_key = st.secrets["AZURE_OPENAI_API_KEY"]
        azure_api_version = st.secrets["API_VERSION"]

        # Validate credentials
        if not all([azure_openai_endpoint, azure_openai_key, azure_api_version]):
            st.error("Missing required secrets. Please check your .streamlit/secrets.toml file.")
            return None

        # Initialize Azure OpenAI client
        client = AzureOpenAI(
            api_key=azure_openai_key,
            api_version=azure_api_version,
            azure_endpoint=azure_openai_endpoint
        )

        deployment_name = "gpt-4o-mini"
        
        prompt = f"""
        Analyze the following SQL stored procedure and return your analysis in JSON format:
        
        {file_content}
        
        Extract and provide:
        1. The name of the stored procedure
        2. The scope/purpose of the stored procedure with details of 4-5 lines.
        3. High-priority optimization opportunities (up to 5), focusing on:
           - Unused temp tables
           - Cursors that can be replaced with CTEs
           - Multiple UPDATE/DELETE statements that can be combined
           - Poor indexing patterns
           - Nested queries with performance issues
           - Any other critical performance issues
        
        For each optimization opportunity, provide:
        - Type of optimization
        - Line number or location in code
        - Existing code snippet (as is)
        - Optimized code snippet (your suggestion)
        - Brief explanation of the benefit
        
        Structure your response as valid JSON that matches this format exactly:
        ```json
        {{
            "procedure_name": "name_here",
            "scope": "description_here",
            "optimizations": [
                {{
                    "type": "type of optimization",
                    "line_number": "approximate line number or range",
                    "existing_logic": "current code snippet",
                    "optimized_logic": "improved code snippet",
                    "explanation": "brief explanation"
                }}
            ]
        }}
        ```
        
        Ensure your response is properly formatted JSON and nothing else.
        """

        response = client.chat.completions.create(
            model=deployment_name,
            messages=[
                {"role": "system", "content": "You are an expert SQL database optimizer that always returns responses in valid JSON format."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            response_format={"type": "json_object"}  # Explicitly request JSON response
        )
        
        # Extract the JSON from the response
        analysis_result = response.choices[0].message.content
        
        # Debug: Display raw response for troubleshooting
        st.sidebar.expander("Debug Raw Response", expanded=False).code(analysis_result)
        
        # Clean the response: Remove any markdown formatting if present
        cleaned_response = analysis_result.strip()
        if cleaned_response.startswith("```json"):
            cleaned_response = cleaned_response[7:]  # Remove ```json prefix
        if cleaned_response.endswith("```"):
            cleaned_response = cleaned_response[:-3]  # Remove ``` suffix
            
        # Parse the JSON
        try:
            return json.loads(cleaned_response)
        except json.JSONDecodeError as e:
            st.error(f"Failed to parse JSON response: {str(e)}")
            st.code(cleaned_response)  # Show the problematic response
            return None
    
    except Exception as e:
        st.error(f"Error during analysis: {str(e)}")
        import traceback
        st.sidebar.expander("Error Details", expanded=False).code(traceback.format_exc())
        return None

# UI Components
st.title("SQL Stored Procedure Analyzer")
st.write("Upload a SQL stored procedure file for AI-powered optimization analysis")

# # Sidebar for configuration information
# with st.sidebar:
#     st.header("Configuration")
#     st.info("This app uses Azure OpenAI configured through environment variables.")
    
#     # Add a configuration check section
#     config_expander = st.expander("Check Configuration", expanded=False)
#     with config_expander:
#         # Load environment variables to check if they're set
#         load_dotenv()
        
#         azure_openai_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")  
#         azure_openai_key = os.getenv("AZURE_OPENAI_KEY")
#         azure_api_version = os.getenv("API_VERSION")
        
#         # Display status of environment variables
#         st.write("Environment Variables Status:")
#         st.write(f"AZURE_OPENAI_ENDPOINT: {'✅ Set' if azure_openai_endpoint else '❌ Missing'}")
#         st.write(f"AZURE_OPENAI_KEY: {'✅ Set' if azure_openai_key else '❌ Missing'}")
#         st.write(f"API_VERSION: {'✅ Set' if azure_api_version else '❌ Missing'}")
        
#         if not all([azure_openai_endpoint, azure_openai_key, azure_api_version]):
#             st.warning("Some required environment variables are missing. Please set them in your .env file.")
            
#             # Option to set variables manually for testing
#             st.subheader("Temporary Setup (Session Only)")
#             temp_endpoint = st.text_input("Azure OpenAI Endpoint", value=azure_openai_endpoint or "")
#             temp_key = st.text_input("Azure OpenAI Key", value="", type="password")
#             temp_api_version = st.text_input("API Version", value=azure_api_version or "2023-05-15")
            
#             if st.button("Apply Temporary Settings"):
#                 os.environ["AZURE_OPENAI_ENDPOINT"] = temp_endpoint
#                 os.environ["AZURE_OPENAI_KEY"] = temp_key
#                 os.environ["API_VERSION"] = temp_api_version
#                 st.success("Temporary settings applied for this session!")
    
st.markdown("---")
st.markdown("""
    ### About This Tool
    This app analyzes SQL stored procedures using AI to identify optimization opportunities.
    
    **Features:**
    - Extract procedure name and purpose
    - Identify optimization opportunities
    - Generate improved SQL code
    - Provide a summary of changes
    - Download formatted report as Word document
    """)

# Sample SQL button for testing
if st.button("Load Sample SQL for Testing"):
    sample_sql = """
    CREATE PROCEDURE usp_GetCustomerOrders
    @CustomerId INT
    AS
    BEGIN
        SET NOCOUNT ON;
        
        -- Create temp table to store order data
        CREATE TABLE #TempOrders (
            OrderId INT,
            OrderDate DATETIME,
            OrderAmount DECIMAL(18,2)
        )
        
        -- Insert data into temp table
        INSERT INTO #TempOrders
        SELECT 
            OrderId,
            OrderDate,
            OrderAmount
        FROM Orders
        WHERE CustomerId = @CustomerId
        
        -- Cursor to process orders
        DECLARE @OrderId INT
        DECLARE @OrderDate DATETIME
        
        DECLARE order_cursor CURSOR FOR
        SELECT OrderId, OrderDate FROM #TempOrders
        
        OPEN order_cursor
        FETCH NEXT FROM order_cursor INTO @OrderId, @OrderDate
        
        WHILE @@FETCH_STATUS = 0
        BEGIN
            -- Update order status
            UPDATE Orders SET Status = 'Processed' WHERE OrderId = @OrderId
            UPDATE Orders SET LastModified = GETDATE() WHERE OrderId = @OrderId
            
            -- Process order details
            UPDATE OrderDetails 
            SET Processed = 1 
            WHERE OrderId = @OrderId
            
            FETCH NEXT FROM order_cursor INTO @OrderId, @OrderDate
        END
        
        CLOSE order_cursor
        DEALLOCATE order_cursor
        
        -- Return results
        SELECT 
            c.CustomerName,
            o.OrderId,
            o.OrderDate,
            o.OrderAmount,
            (SELECT COUNT(*) FROM OrderDetails WHERE OrderId = o.OrderId) AS ItemCount
        FROM 
            Customers c
            INNER JOIN Orders o ON c.CustomerId = o.CustomerId
        WHERE 
            c.CustomerId = @CustomerId
            
        -- Cleanup
        DROP TABLE #TempOrders
    END
    """
    st.session_state['sample_sql'] = sample_sql
    st.success("Sample SQL loaded! Click 'Analyze' to process it.")

# File upload component
uploaded_file = st.file_uploader("Upload SQL Stored Procedure", type=["sql"])

# Get SQL either from upload or sample
sql_content = None
if uploaded_file:
    sql_content = uploaded_file.getvalue().decode("utf-8")
elif 'sample_sql' in st.session_state:
    sql_content = st.session_state['sample_sql']
    st.info("Using sample SQL procedure. You can upload your own file to replace it.")

if sql_content:
    # Display the SQL
    with st.expander("View SQL Code", expanded=False):
        st.code(sql_content, language="sql")
    
    # Analysis button
    if st.button("Analyze SQL Procedure"):
        # Run analysis
        with st.spinner("Analyzing stored procedure... This may take up to 30 seconds."):
            analysis = analyze_stored_procedure(sql_content)
        
        if analysis:
            # Display results in tabs
            tab1, tab2 = st.tabs(["Analysis", "Download Report"])
            
            with tab1:
                # Display the procedure name and scope
                st.header(f"🔹 Stored Proc Name: `{analysis['procedure_name']}`")
                
                st.subheader("🔹 Scope:")
                st.write(analysis["scope"])
                
                # Display optimization steps
                st.subheader("🔹 Optimization Steps:")
                
                for i, opt in enumerate(analysis["optimizations"], 1):
                    st.markdown(f"⚙️ **Step {i}**: {opt['type']}")
                    
                    st.markdown("**Existing Logic:**")
                    st.code(opt["existing_logic"], language="sql")
                    
                    st.markdown("**Optimized Logic:**")
                    st.code(opt["optimized_logic"], language="sql")
                    
                    st.markdown(f"*{opt['explanation']}*")
                    st.markdown("---")
                
                # Create and display summary table
                st.subheader("🔹 Summary:")
                
                summary_data = []
                for opt in analysis["optimizations"]:
                    summary_data.append({
                        "Type of Change": opt["type"],
                        "Line Number": opt["line_number"],
                        "Original Code Snippet": opt["existing_logic"][:40] + "..." if len(opt["existing_logic"]) > 40 else opt["existing_logic"],
                        "Optimized Code Snippet": opt["optimized_logic"][:40] + "..." if len(opt["optimized_logic"]) > 40 else opt["optimized_logic"],
                        "Optimization Explanation": opt["explanation"]
                    })
                
                summary_df = pd.DataFrame(summary_data)
                
                # Display as a formatted table with custom styling
                st.markdown("""
                <style>
                .summary-table {
                    font-size: 0.85rem;
                    border-collapse: collapse;
                    width: 100%;
                }
                .summary-table th {
                    background-color: #f2f2f2;
                    text-align: left;
                    padding: 8px;
                    border: 1px solid #ddd;
                }
                .summary-table td {
                    text-align: left;
                    padding: 8px;
                    border: 1px solid #ddd;
                }
                .summary-table tr:nth-child(even) {
                    background-color: #f9f9f9;
                }
                </style>
                """, unsafe_allow_html=True)
                
                # Convert dataframe to HTML table with custom classes
                table_html = summary_df.to_html(classes='summary-table', escape=False, index=False)
                st.markdown(table_html, unsafe_allow_html=True)
            
            with tab2:
                # Create Word document
                with st.spinner("Generating Word document..."):
                    docx_bytes = create_word_document(analysis)
                
                # Provide download button for DOCX
                st.download_button(
                    label="⬇️ Download Report as Word Document",
                    data=docx_bytes,
                    file_name=f"{analysis['procedure_name']}_analysis.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key='docx-download'
                )
                
                # Also provide markdown option
                report_md = f"""# SQL Stored Procedure Analysis Report

## Procedure Name: {analysis['procedure_name']}

## Scope:
{analysis['scope']}

## Optimization Steps:
"""
                
                for i, opt in enumerate(analysis["optimizations"], 1):
                    report_md += f"""
### Step {i}: {opt['type']}

**Existing Logic:**
```sql
{opt['existing_logic']}
```

**Optimized Logic:**
```sql
{opt['optimized_logic']}
```

*{opt['explanation']}*

---
"""
                
                report_md += "\n## Summary Table:\n\n"
                report_md += summary_df.to_markdown(index=False)
                
                st.download_button(
                    label="⬇️ Download Report as Markdown",
                    data=report_md,
                    file_name=f"{analysis['procedure_name']}_analysis.md",
                    mime="text/markdown",
                    key='md-download'
                )
                
                st.info("The Word document (.docx) contains the same content as shown in the 'Analysis' tab, but in a properly formatted document for sharing.")
        else:
            st.error("Analysis could not be completed. Please check the Debug section in the sidebar for more details.")

else:
    # Show example when no file is uploaded
    st.info("Please upload a SQL stored procedure file (.sql) or use the sample SQL to begin analysis.")
    
    with st.expander("See Example Analysis"):
        st.markdown("""
        ## Example Output
        
        🔹 **Stored Proc Name:** `usp_get_customer_data`
        
        🔹 **Scope:**  
        This procedure retrieves customer data including their purchase history, last login time, and calculates their loyalty score using internal metrics.
        
        🔹 **Optimization Steps:**
        
        ⚙️ **Step 1:** Replace Multiple Updates
        
        **Existing Logic:**
        ```sql
        UPDATE table SET col1 = val WHERE condition;
        UPDATE table SET col2 = val WHERE condition;
        ```
        
        **Optimized Logic:**
        ```sql
        UPDATE table 
        SET col1 = val, 
            col2 = val 
        WHERE condition;
        ```
        
        *Reduces write operations and improves efficiency.*
        """)
        
        # Example of the summary table
        st.markdown("🔹 **Summary:**")
        
        example_data = [{
            "Type of Change": "Replace Multiple Updates",
            "Line Number": "Identified in multiple places",
            "Original Code Snippet": "UPDATE col1 + UPDATE col2",
            "Optimized Code Snippet": "UPDATE col1, col2",
            "Optimization Explanation": "Reduces write operations and improves efficiency."
        }, {
            "Type of Change": "Index on Temp Tables",
            "Line Number": "Where temp tables are created",
            "Original Code Snippet": "CREATE TABLE #temp",
            "Optimized Code Snippet": "CREATE INDEX ON #temp",
            "Optimization Explanation": "Improves performance by speeding up lookups and joins."
        }]
        
        example_df = pd.DataFrame(example_data)
        
        # Display example table with styling
        st.markdown("""
        <style>
        .summary-table {
            font-size: 0.85rem;
            border-collapse: collapse;
            width: 100%;
        }
        .summary-table th {
            background-color: #f2f2f2;
            text-align: left;
            padding: 8px;
            border: 1px solid #ddd;
        }
        .summary-table td {
            text-align: left;
            padding: 8px;
            border: 1px solid #ddd;
        }
        .summary-table tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        </style>
        """, unsafe_allow_html=True)
        
        table_html = example_df.to_html(classes='summary-table', escape=False, index=False)
        st.markdown(table_html, unsafe_allow_html=True)

# Add footer
st.markdown("---")
st.caption("SQL Stored Procedure Analyzer powered by Azure OpenAI")