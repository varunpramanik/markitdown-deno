import { loadPyodide } from 'npm:pyodide';
import { join as joinPaths, SEPARATOR as SEP } from "jsr:@std/path";



const INPUT_FILE = 'example.xlsx';
const OUTPUT_FILE = INPUT_FILE + '.md';

const INPUT_PATH = joinPaths(Deno.cwd(), INPUT_FILE);
const OUTPUT_PATH = joinPaths(Deno.cwd(), OUTPUT_FILE);


// Load and set up Pyodide
//
//
const pyodide = await loadPyodide();
await pyodide.loadPackage(['micropip']);

// Run conversion
//
//
convertToMarkdown(INPUT_PATH, OUTPUT_PATH);


/**
 * 
 * 
 * 
 */
async function convertToMarkdown(inputPath: string, outputPath: string): Promise<void> {

    // Read the data from the filesystem
    const inputData = Deno.readFileSync(inputPath);

    // Set up and write the input to Pyodide's virtual filesystem
    const inputFileName = inputPath.split(SEP).pop() as string;
    const outputFileName = outputPath.split(SEP).pop() as string;
    const virtualInput = '/' + inputFileName;
    const virtualOutput = '/' + outputFileName;
    pyodide.FS.writeFile(virtualInput, inputData);

    // Execute the Python code

    const pythonCode = `
        import micropip
        await micropip.install("markitdown")
        from markitdown import MarkItDown
        
        markitdown = MarkItDown()
        markitdown.convert(source="${virtualInput}", output_path="${virtualOutput}")
        `;

    await pyodide.runPythonAsync(pythonCode);

    // Retrieve and save the output data
    const outputData = pyodide.FS.readFile(virtualOutput);
    Deno.writeFileSync(outputPath, outputData);

    console.log(`Output written to ${outputPath}`);
}