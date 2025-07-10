# Copilot Instructions

## Purpose

This document provides structured guidelines for integrating AI project workflows in VSCode, ensuring maximum automation, reliability, and traceability across conversations and executions. It is based on the "AI Project Workflow Guide" and tailored for the Reference_project. The workflow assumes a Machine Code Playground (MCP) environment with plugins like Sequential Thinking, REPL, Fetch, Puppeteer, GitHub, and others.

## Activation & Setup

- **Automatic Activation**: All tools and instructions are active by default in every conversation for the project. Use them proactively without waiting for explicit user prompts.
- **Environment Checks**: Before starting any task, ensure the required MCP instance is running. If it's stopped, start it now. This guarantees all tools and workflows operate correctly.
- **Default Active Tools**:
  - Kite: Trading and stock market data
  - Git/GitHub: Version control and repository integration
  - Context7: Enhanced contextual reasoning and memory management
  - Playwright/Puppeteer: Web scraping and browser automation
  - Excel: Spreadsheet data handling and calculations
  - Filesystem/Desktop Commander: Local file and system operations
  - Memory: Persistent storage of conversation context
  - Sequential Thinking: Breaking down complex problems
  - Fetch/DuckDuckGo/Tavily: Web search and HTTP content retrieval
  - Zoom: Meeting scheduling and interaction
  - Notion: Document management and collaboration
  - Everything-Search: Fast local file search
  - Time: Time and timezone conversions
  - WhatsApp-MCP: Access to personal WhatsApp (messages, media)
  - PDF-Reader: Text extraction from PDF documents
  - Artifacts: Exporting code, visualizations, or documents

## Default Workflow Steps

1. **Context Retrieval**: Begin by retrieving and reviewing the last 10 minutes of conversation history (if available) to establish context. This ensures continuity in multi-turn interactions.
2. **Sequential Thinking**: Immediately use Sequential Thinking to break down the user's request into core components. Deconstruct the problem, identify concepts and dependencies, and select appropriate tools.
3. **Search & Research**: Conduct broad searches using Tavily or DuckDuckGo for initial information. Refine queries with specific terms or filters and log the query strings, result counts, and access times.
4. **Web Verification**: For top search results, use Playwright or Puppeteer to visit websites. Take annotated screenshots, log navigation steps, and extract data (e.g., fill forms, parse tables). Verify that expected content appears before proceeding.
5. **Data Processing**: Use REPL, Excel, or scripts for data tasks. Load any structured data (CSV, JSON) and perform calculations or analysis. Generate tables, charts, or summaries of insights. Link results to memory or the knowledge graph if useful.
6. **Synthesis & Output**: Combine insights, data, code, and visuals into a coherent final product. Create well-documented outputs (Markdown notes, Jupyter notebooks, Python scripts, CSV files, etc.). Highlight key findings, decision logic, and ensure traceability to sources.
7. **Memory Update**: After completing a task, check if any new entities or facts should be added to persistent memory. Update memory so the information is available for future context.
8. **Finalization**: Always monitor command outputs for success. Use PowerShell (preferred over plain CMD) for scripting tasks. At the end of each task, commit changes to GitHub with a clear commit message. If any files are too large, mention them and add them to .gitignore instead of pushing.

## Tools and Usage

Each tool's role is described in simple, actionable language:

- **Sequential Thinking**: Decompose multi-step problems into clear, ordered steps. Write out the reasoning explicitly and consider alternatives if needed.
- **Context7**: Keep track of conversation context and long-term project knowledge. Use it to manage and recall prior information.
- **Memory**: Store new entities or facts that should persist across conversations. Check and update memory after each task to maintain context.
- **Kite**: Retrieve up-to-date trading and stock market data when financial information is needed.
- **Git/GitHub**: Track code changes, commit results, and synchronize with the remote repository. Exclude large files by adding them to .gitignore.
- **Fetch**: Perform HTTP requests to retrieve web content or call APIs. Use it for structured data or specific endpoints.
- **DuckDuckGo**: Conduct general web searches in a privacy-focused way. Use broad queries for initial research.
- **Tavily**: Use this AI-powered search engine for comprehensive, contextual web information retrieval.
- **Excel**: Work with spreadsheet data and complex formulas. Use Excel for tabular data analysis or when formulas are required.
- **Playwright/Puppeteer**: Automate web browser interactions. Use these for site navigation, form filling, screenshot capture, and data scraping.
- **Filesystem**: Navigate and manage files and directories on the local system. Create, read, update, or delete files as needed.
- **Desktop Commander**: Perform advanced desktop or system file operations beyond the basic filesystem commands.
- **Artifacts**: Export final outputs (code files, data files, images, documents) to share outside the chat. Use artifacts for anything longer than a message or requiring download.
- **Zoom**: Schedule or join meetings and retrieve relevant meeting information.
- **Notion**: Access or update project documentation in Notion for collaboration.
- **Everything-Search**: Quickly search for files on the local machine by name or content.
- **Time**: Convert between time zones or get the current time in different zones.
- **WhatsApp-MCP**: Search and read personal WhatsApp conversations, including text, images, and media.
- **PDF-Reader**: Search within and extract text from PDF documents to obtain structured information.
- **REPL/Analysis Tools**: Run scripts or code snippets (e.g., Python, PowerShell) for data computation or analysis. Use REPL for repeatable computations and experiments.

## Core Workflow Phases

1. **Initial Analysis**

   - Use Sequential Thinking to break down the user's request step-by-step
   - Identify key concepts, relationships, and dependencies in the problem
   - Decide which tools are needed for each part of the task
   - Plan tasks that can run in parallel when possible

2. **Search Phase**

   - Start with broad, semantic queries using Tavily or DuckDuckGo
   - Use filters and offsets to narrow results as needed
   - Log each query string, the number of results, and search timestamps
   - Save summary of search results, including titles and snippets

3. **Deep Verification**

   - Use Playwright or Puppeteer to visit top search result websites
   - Capture annotated screenshots of important pages
   - Log every navigation step (URLs visited)
   - Interact with pages: fill forms, click links, and extract structured data
   - Confirm that expected content appears before moving on

4. **Data Processing**

   - Load structured data (CSV, JSON) into REPL, Excel, or scripts
   - Perform calculations, filtering, or transformations
   - Generate analysis outputs: tables, charts, or statistics
   - Document all code and formulas used for transparency
   - Optionally save results to memory or a knowledge graph

5. **Synthesis & Output**

   - Combine data, code, insights, and visuals into final artifacts
   - Create well-documented output files (Markdown, Jupyter Notebook, Python scripts, etc.)
   - Explain final insights and show how decisions were reached
   - Ensure outputs highlight sources and maintain traceability to original data

6. **Memory Update**
   - After completing tasks, determine if new information should be stored in Memory
   - Add new entities or facts to persistent memory to aid future context
   - Ensure memory links back to conversation context or source when appropriate

## Tool-Specific Best Practices

- **Tavily**:

  - Use for comprehensive web searches and contextual queries
  - It can handle complex, multi-faceted searches better than standard queries

- **Fetch/DuckDuckGo**:

  - Specify count and offset parameters to control the number of search results
  - Document each query string and a brief summary of results
  - Save metadata: webpage titles, URLs, snippets, and access timestamps for each result

- **Puppeteer/Playwright**:

  - Always check that navigation was successful (no errors) before extracting data
  - Use precise selectors to scrape the needed information
  - Take and annotate screenshots for evidence of important pages or results
  - Handle page load errors gracefully by retrying or logging the issue

- **Sequential Thinking**:

  - Write out each step of your reasoning as you break down the problem
  - If there are multiple ways to solve a problem, consider and document alternative approaches
  - Allow for reevaluating steps in the middle of a task if new information appears

- **REPL/Analysis Tools**:

  - Use code (Python, scripts, or Excel formulas) to perform repeatable data computations
  - Show all code, functions, or formulas used so the logic is transparent
  - Annotate results (for example, label chart axes or explain table columns) to make them easy to review

- **Artifacts**:
  - Use the Artifacts tool to output results that are too long for a chat message or need to be saved (e.g., large tables, code files, images)
  - Clearly annotate the purpose and structure of each artifact
  - Link artifacts back to the original conversation or data source when relevant

## General Best Practices

- **Proactive Tool Use**: Activate tools based on the task requirements, not just when prompted. If a step involves data lookup or computation, use the appropriate tool immediately.
- **Parallel Execution**: Run independent tasks or tool commands in parallel when possible to improve efficiency.
- **Terminal Monitoring**: Always check terminal output after running commands and proceed to the next step without waiting. Don't wait indefinitely for processes that may have completed or stalled.
- **Problems Validation**: Always check the Problems tab before pushing code to GitHub to identify and fix any errors or warnings.
- **Documentation-First**: Log or cite your work at every major step. Keep a detailed record of actions, results, and decisions.
- **Knowledge Retention**: Store important insights or data in memory or a knowledge graph. This prevents repeating work in future conversations.
- **Full Pipeline Triggering**: For multi-step tasks, the entire workflow (from analysis to output) should be triggered automatically, without waiting for user-by-user prompts at each step.
