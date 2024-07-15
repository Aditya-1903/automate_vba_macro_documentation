import streamlit as st
import os
import re
from oletools.olevba import VBA_Parser
from langchain_core.prompts import PromptTemplate
from langchain_groq import ChatGroq

llm = ChatGroq(
        temperature=0,
        model="mixtral-8x7b-32768",
        api_key=os.environ['GROQ_API_KEY']  # Replace with your actual API key
    )

# Ensure necessary directories exist
os.makedirs("uploads", exist_ok=True)
os.makedirs("outputs", exist_ok=True)
os.makedirs("vba", exist_ok=True)

def extract_vba_from_excel(file_path, output_dir):
    vba_parser = VBA_Parser(file_path)
    if vba_parser.detect_vba_macros():
        vba_modules = vba_parser.extract_all_macros()
    else:
        st.write("No Macros found in the Given Workbook!")
        return ""

    code = ""
    for _, _, vba_name, vba_code in vba_modules:
        if vba_name.endswith('.bas'):
            code += vba_code

    if code:
        code = re.sub(r"\s+", " ", code).strip()
        output_file_path = os.path.join(output_dir, 'vba_code.txt')
        with open(output_file_path, 'w') as f:
            f.write(code)
        return code
    else:
        st.write("The Macros do not contain any written code!")
        return ""

def analyze_vba(vba_code_path):
    with open(vba_code_path, 'r') as file:
        vba_code = file.read()


    prompt_template_vba_macro_documentation = """
    You are a helpful assistant.
    You are responsible for analyzing the given VBA Macro code and generating a comprehensive documentation of the underlying logic, data flow and process flow.
    Your response must contain the underlying logic, data flow and process flow for each of the subroutines available in the given code.
    VBA Code:{question}
    """

    prompt = PromptTemplate(template=prompt_template_vba_macro_documentation)
    query_with_prompt = prompt.format(question=vba_code)
    vba_macro_documentation = llm.invoke(query_with_prompt).content.replace('\n', ' ')

    prompt_template_functional_logic = """
    You are a helpful assistant.
    You are responsible for extracting and explaining the functional logic embedded within the given VBA macro.
    Your response should help technical and non-technical stakeholders understand the business logic, supporting better decision-making
    and transformation efforts.
    VBA Code:{question}
    """

    prompt = PromptTemplate(template=prompt_template_functional_logic)
    query_with_prompt = prompt.format(question=vba_code)
    vba_macro_functional_logic = llm.invoke(query_with_prompt).content.replace('\n', ' ')

    with open('outputs/vba_macro_documentation.txt', 'w') as doc_file:
        doc_file.write(vba_macro_documentation)

    with open('outputs/vba_macro_functional_logic.txt', 'w') as logic_file:
        logic_file.write(vba_macro_functional_logic)

    return vba_macro_documentation, vba_macro_functional_logic

def analyze_code_quality(vba_code):
    

    prompt_template_code_quality = """
    Please analyze the provided VBA Macro code and evaluate its quality and efficiency.
    Identify potential inefficiencies, redundant code, and optimization opportunities to improve macro performance and reliability.
    Pay particular attention in analyzing:
      Code Efficiency: Assess if the code is optimized for speed and resource usage.
      Redundant Code: Identify any sections of code that repeat functionality unnecessarily.
      Optimization Opportunities: Suggest improvements or optimizations that could enhance the macro's performance.
    Please provide only a detailed feedback on each of these aspects to guide improvements in the given VBA Macro code.
    Do not provide any VBA code as a part of the response.
    Here is the VBA Macro code snippet:
    VBA Code:{question}
    """

    prompt = PromptTemplate(template=prompt_template_code_quality)
    query_with_prompt = prompt.format(question=vba_code)
    vba_macro_code_quality = llm.invoke(query_with_prompt).content.replace('\n', ' ')

    output_path = os.path.join('outputs', 'vba_macro_code_quality.txt')
    with open(output_path, 'w') as quality_file:
        quality_file.write(vba_macro_code_quality)

    return vba_macro_code_quality

def analyze_data_flow(vba_code):
    

    prompt_template_data_flow = """
    Please analyze the provided VBA Macro code and evaluate the data flow within the macros.
    Identify bottlenecks and opportunities for optimization to enhance efficiency and performance in data processing tasks. Specifically, focus on:
    Data Flow Analysis: Map out how data moves through the macro from input to output.
    Bottlenecks: Identify points in the code where data processing slows down or encounters inefficiencies.
    Optimization Opportunities: Recommend changes or optimizations to streamline data processing and improve overall performance.
    Resource Usage: Assess how efficiently resources are utilized during data processing.
    Scalability: Consider how well the macro handles varying data volumes and complexity.

    Please provide detailed insights and recommendations to optimize the VBA Macro code for improved efficiency and performance in data processing tasks.
    Here is the VBA Macro code snippet:
    VBA Code:{question}
    """

    prompt = PromptTemplate(template=prompt_template_data_flow)
    query_with_prompt = prompt.format(question=vba_code)
    vba_macro_data_flow = llm.invoke(query_with_prompt).content.replace('\n', ' ')

    output_path = os.path.join('outputs', 'vba_macro_data_flow.txt')
    with open(output_path, 'w') as data_flow_file:
        data_flow_file.write(vba_macro_data_flow)

    return vba_macro_data_flow

def format_vba_content(content):
    paragraphs = content.split('  ')
    formatted_content = ""
    
    for paragraph in paragraphs:
        paragraph = paragraph.strip()
        if paragraph.startswith('Functional Logic:'):
            formatted_content += f"**{paragraph}**\n\n"
        elif paragraph.startswith('-'):
            formatted_content += f"<ul><li>{paragraph[1:].strip()}</li></ul>\n\n"
        elif paragraph.startswith('~'):
            formatted_content += f"<h4>{paragraph}</h4>\n\n"
        else:
            formatted_content += f"{paragraph}\n\n"

    return formatted_content

def extract_nodes_and_links(vba_code):
    nodes = []
    links = []

    # Regex patterns to match subroutine names and calls
    subroutine_pattern = r'Sub\s+(\w+)\('
    call_pattern = r'Call\s+(\w+)\b'

    # Find all subroutine names (nodes)
    subroutine_matches = re.findall(subroutine_pattern, vba_code)
    nodes.extend(subroutine_matches)

    # Find all calls between subroutines (links)
    for match in re.finditer(subroutine_pattern, vba_code):
        subroutine_name = match.group(1)
        start = match.end()
        end = re.search(r'End\s+Sub', vba_code[start:]).start()
        subroutine_body = vba_code[start:start+end]
        call_matches = re.findall(call_pattern, subroutine_body)
        for call in call_matches:
            if call in nodes:
                links.append((subroutine_name, call))

    return nodes, links

def update_data_js(vba_code):
    nodes, links = extract_nodes_and_links(vba_code)

    nodes_js = ",\n".join([f"{{ id: '{node}' }}" for node in nodes])
    links_js = ",\n".join([f"{{ source: '{link[0]}', target: '{link[1]}' }}" for link in links])

    data_js_content = f"const nodes = [\n{nodes_js}\n];\n\nconst links = [\n{links_js}\n];\n"

    with open('outputs/data.js', 'w') as f:
        f.write(data_js_content)

    return data_js_content




def refactor_vba(vba_code_path):
    with open(vba_code_path, 'r') as file:
        vba_code = file.read()



    prompt_template_refactor = """
    Please provide recommendations for refactoring or rewriting the given VBA Macro code (only if necessary) using modern programming languages and technologies to enhance maintainability, scalability, and performance.
    Consider the following:
      Language and Technology Recommendations: Suggest suitable modern programming languages (e.g., Python) and technologies (e.g., APIs) that align with the macro's functionality.
      Refactoring Strategies: Recommend specific refactoring techniques to improve code structure, readability, and maintainability.
      Integration Possibilities: Explore integration options with other systems or platforms for enhanced functionality and interoperability.
      Performance Optimization: Identify opportunities to optimize code performance and efficiency in the new environment.
      Compatibility Considerations: Address compatibility issues or considerations when transitioning from VBA to modern languages or technologies.

    Please provide detailed guidance on how to effectively modernize the VBA Macro code, ensuring it meets current industry standards and best practices.

    Here is the VBA Macro code snippet:
    VBA Code:{question}
    """

    prompt = PromptTemplate(template=prompt_template_refactor)
    query_with_prompt = prompt.format(question=vba_code)
    vba_macro_refactor = llm.invoke(query_with_prompt).content
    
    output_file_path = os.path.join('outputs', 'vba_macro_refactor.txt')
    with open(output_file_path, 'w') as f:
        f.write(vba_macro_refactor)
    
    return vba_macro_refactor

def check_vba_security(vba_code_path):
    with open(vba_code_path, 'r') as file:
        vba_code = file.read()

    risky_patterns = [
        r'\bShell\b',
        r'\bExecute\b',
        r'\bActiveX\b'
    ]

    risks_found = []
    for pattern in risky_patterns:
        matches = re.findall(pattern, vba_code, re.IGNORECASE)
        if matches:
            risks_found.extend(matches)

    error_handling = "On Error" in vba_code

    sql_patterns = [
        r'INSERT INTO',
        r'SELECT \* FROM',
        r'UPDATE .* SET',
        r'DELETE FROM'
    ]

    sanitization_issues = []
    for pattern in sql_patterns:
        if re.search(pattern, vba_code, re.IGNORECASE):
            if "'" in vba_code:
                sanitization_issues.append(pattern)

    report = []

    if risks_found:
        report.append("Risky Patterns Found:")
        report.extend([f" - {func}" for func in risks_found])
        report.append("The above listed functions pose a security risk. They may execute external applications, which could be exploited by attackers.\n")
    else:
        report.append("No risky patterns identified in the given VBA Macro.\n")

    if not error_handling:
        report.append("Error Handling: \nNot present. Proper error handling is crucial to prevent the system from exposing sensitive information during errors and to ensure that the application behaves securely and predictably.\n")

    if sanitization_issues:
        report.append("Sanitization Issues Found:")
        report.extend([f" - {issue}" for issue in sanitization_issues])
        report.append("The above listed SQL patterns may be vulnerable to SQL injection if not properly parameterized. Ensure that SQL queries are constructed using parameterized queries or stored procedures to prevent injection attacks.\n")
    else:
        report.append("All user inputs are properly sanitized, reducing the risk of injection attacks.\n")

    output_file_path = os.path.join('outputs', 'vba_security_report.txt')
    with open(output_file_path, 'w') as f:
        f.write('\n'.join(report))
    
    return '\n'.join(report)

# Main Page
def main():
    st.title("Automating VBA Macro Documentation and Transformation")
    st.write("""
    This tool allows you to upload excel workbooks and automate VBA Macro Documentation.
    
    ### Features:
    - **File Upload**: Accepts `.xls` and `.xlsm` files.
    - **VBA Code Extraction**: Extracts VBA code and saves it as a text file.
    - **VBA Code Analysis**: Analyzes the extracted VBA code and categorizes the elements.
    
    Use the sidebar to navigate between different use cases.
    """)

    uploaded_file = st.file_uploader("Upload a .xls or .xlsm file", type=["xls", "xlsm"])
    
    if uploaded_file is not None:
        file_path = os.path.join("uploads", uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.session_state.file_path = file_path  
        st.session_state.file_name = uploaded_file.name  

# Use Case 1 Page
@st.cache_data
@st.cache_resource
def use_case1():
    st.title("Use Case 1: VBA Macro Documentation")

    if 'file_path' not in st.session_state:
        st.write("Please upload a file on the Home page.")
        return
    
    st.write(f"Uploaded File: {st.session_state.file_name}")

    vba_code = extract_vba_from_excel(st.session_state.file_path, "vba")
    if vba_code:
        vba_code_path = os.path.join("vba", "vba_code.txt")
        vba_macro_documentation, _ = analyze_vba(vba_code_path)
        st.subheader("VBA Macro Documentation")
        formatted_doc = format_vba_content(vba_macro_documentation)
        st.markdown(formatted_doc, unsafe_allow_html=True)

        documentation_file = 'outputs/vba_macro_documentation.txt'
        with open(documentation_file, 'r') as file:
            documentation_content = file.read()
        st.download_button(label="Download VBA Macro Documentation", data=documentation_content, file_name='vba_macro_documentation.txt')

    if st.button("Return to Home"):
        st.session_state.page = "Home"
        st.experimental_rerun()

# Use Case 2 Page
def use_case2():
    st.title("Use Case 2: Functional Logic Extractor")

    if 'file_path' not in st.session_state:
        st.write("Please upload a file on the Home page.")
        return

    st.write(f"Uploaded File: {st.session_state.file_name}")  

    vba_code = extract_vba_from_excel(st.session_state.file_path, "vba")
    if vba_code:
        vba_code_path = os.path.join("vba", "vba_code.txt")
        _, vba_macro_functional_logic = analyze_vba(vba_code_path)
        st.subheader("Functional Logic Extractor")
        if vba_macro_functional_logic:
            formatted_logic = format_vba_content(vba_macro_functional_logic)
            st.markdown(formatted_logic, unsafe_allow_html=True)

            functional_logic_file = 'outputs/vba_macro_functional_logic.txt'
            with open(functional_logic_file, 'r') as file:
                functional_logic_content = file.read()
            st.download_button(label="Download VBA Macro Functional Logic", data=functional_logic_content, file_name='vba_macro_functional_logic.txt')
        else:
            st.write("No functional logic found in the VBA code.")

    if st.button("Return to Home"):
        st.session_state.page = "Home"
        st.experimental_rerun()



# Use Case 3 Page
def use_case3():
    st.title("Use Case 3: Process Flow Visualization")

    if 'file_path' not in st.session_state:
        st.write("Please upload a file on the Home page.")
        return

    st.write(f"Uploaded File: {st.session_state.file_name}")

    with open("outputs/flow_diagram.html", "r") as html_file:
        st.components.v1.html(html_file.read(), height=800)

# Use Case 4 Page
def use_case4():
    st.title("Use Case 4: Code Quality and Efficiency Analyzer")

    if 'file_path' not in st.session_state:
        st.write("Please upload a file on the Home page.")
        return

    st.write(f"Uploaded File: {st.session_state.file_name}")  

    vba_code = extract_vba_from_excel(st.session_state.file_path, "vba")
    if vba_code:
        vba_macro_code_quality = analyze_code_quality(vba_code)
        st.subheader("VBA Macro Code Quality")
        if vba_macro_code_quality:
            formatted_quality = format_vba_content(vba_macro_code_quality)
            st.markdown(formatted_quality, unsafe_allow_html=True)

            code_quality_file = 'outputs/vba_macro_code_quality.txt'
            with open(code_quality_file, 'r') as file:
                code_quality_content = file.read()
            st.download_button(label="Download VBA Macro Code Quality", data=code_quality_content, file_name='vba_macro_code_quality.txt')
        else:
            st.write("No code quality analysis available.")

    if st.button("Return to Home"):
        st.session_state.page = "Home"
        st.experimental_rerun()

# Use Case 7 Page
def use_case7():
    st.title("Use Case 7: Data Flow Analysis Optimization")

    if 'file_path' not in st.session_state:
        st.write("Please upload a file on the Home page.")
        return

    st.write(f"Uploaded File: {st.session_state.file_name}")

    vba_code = extract_vba_from_excel(st.session_state.file_path, "vba")
    if vba_code:
        vba_macro_data_flow = analyze_data_flow(vba_code)
        st.subheader("VBA Macro Data Flow")
        if vba_macro_data_flow:
            formatted_data_flow = format_vba_content(vba_macro_data_flow)
            st.markdown(formatted_data_flow, unsafe_allow_html=True)

            data_flow_file = 'outputs/vba_macro_data_flow.txt'
            with open(data_flow_file, 'r') as file:
                data_flow_content = file.read()
            st.download_button(label="Download VBA Macro Data Flow", data=data_flow_content, file_name='vba_macro_data_flow.txt')
        else:
            st.write("No data flow analysis available.")

    if st.button("Return to Home"):
        st.session_state.page = "Home"
        st.experimental_rerun()


# Use Case 8 Page
def use_case8():
    st.title("Use Case 8: Legacy Macro Modernization Assistant")

    if 'file_path' not in st.session_state:
        st.write("Please upload a file on the Home page.")
        return

    st.write(f"Uploaded File: {st.session_state.file_name}")

    vba_code = extract_vba_from_excel(st.session_state.file_path, "vba")
    if vba_code:
        vba_macro_refactor = refactor_vba(os.path.join("vba", "vba_code.txt"))
        st.subheader("VBA Macro Refactor")
        if vba_macro_refactor:
            st.markdown(vba_macro_refactor)

            refactor_file = 'outputs/vba_macro_refactor.txt'
            with open(refactor_file, 'r') as file:
                refactor_content = file.read()
            st.download_button(label="Download VBA Macro Refactor", data=refactor_content, file_name='vba_macro_refactor.txt')
        else:
            st.write("No refactor recommendations available.")

    if st.button("Return to Home"):
        st.session_state.page = "Home"
        st.experimental_rerun()

# Use Case 9 Page
def use_case9():
    st.title("Use Case 9: Security and Compliance Checker")

    if 'file_path' not in st.session_state:
        st.write("Please upload a file on the Home page.")
        return

    st.write(f"Uploaded File: {st.session_state.file_name}")

    vba_code = extract_vba_from_excel(st.session_state.file_path, "vba")
    if vba_code:
        security_report = check_vba_security(os.path.join("vba", "vba_code.txt"))
        st.subheader("VBA Macro Security Analysis")
        if security_report:
            st.text(security_report)

            security_file = 'outputs/vba_security_report.txt'
            with open(security_file, 'r') as file:
                security_content = file.read()
            st.download_button(label="Download VBA Macro Security Analysis", data=security_content, file_name='vba_security_report.txt')
        else:
            st.write("No security issues found.")

    if st.button("Return to Home"):
        st.session_state.page = "Home"
        st.experimental_rerun()

st.sidebar.title("Navigation")
if 'page' not in st.session_state:
    st.session_state.page = "Home"

page_options = [
    "Home",
    "VBA Macro Documentation",
    "Functional Logic Extractor",
    "Process Flow Visualization",
    "Code Quality and Efficiency Analyzer",
    "Data Flow Analysis Optimization",
    "Legacy Macro Modernization Assistant",
    "Security and Compliance Checker"
]

page = st.sidebar.selectbox("Choose a page", page_options, index=page_options.index(st.session_state.page))

if page != st.session_state.page:
    st.session_state.page = page

if st.session_state.page == "Home":
    main()
elif st.session_state.page == "VBA Macro Documentation":
    use_case1()
elif st.session_state.page == "Functional Logic Extractor":
    use_case2()
elif st.session_state.page == "Process Flow Visualization":
    use_case3()
elif st.session_state.page == "Code Quality and Efficiency Analyzer":
    use_case4()
elif st.session_state.page == "Data Flow Analysis Optimization":
    use_case7()
elif st.session_state.page == "Legacy Macro Modernization Assistant":
    use_case8()
elif st.session_state.page == "Security and Compliance Checker":
    use_case9()

