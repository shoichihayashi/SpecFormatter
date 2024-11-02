import React, { useState } from 'react'
import axios from 'axios'
import { FileUploader } from "react-drag-drop-files";
import "./styles.css"


const fileTypes = ["docx"]

const Form = () => {
    const [sourceFile, setSource] = useState(null)
    const [targetFiles, setTargets] = useState([])
    const [message, setMessage] = useState('')
    const [downloadLink, setDownloadLink] = useState('')

    const handleSource = (file) => {
        setDownloadLink('')
        setSource(file)
    }

    const handleTarget = (files) => {
        setDownloadLink('')
        const filesArray = Array.isArray(files) ? files : Array.from(files);
        setTargets(filesArray)
    }

    const handleSubmit = async (e) => {
        e.preventDefault()

        const formData = new FormData()

        // single source file
        formData.append('source_file', sourceFile)

        // for loop to append target files to the array
        targetFiles.forEach((file, index) => {
            formData.append(`target_file_${index}`, file)
        })

        try {
            const response = await axios.post('http://localhost:5000/process', formData, {
                responseType: 'blob'
            })

            if (response.data) {
                setMessage(response.data.message)

                const url = window.URL.createObjectURL(new Blob([response.data]))
                setDownloadLink(url)
            } else {
                setMessage('Error processing files')
            }
        } catch(error) {
            setMessage('Error processing the files')
        }
    }

    return (
        <div>
            <h1>SpecFormatter</h1>
            <form onSubmit={handleSubmit}>
                <div>
                    <label>Source document:</label>
                    <FileUploader
                        types={fileTypes}
                        multiple={false}
                        name="sourceFile"
                        handleChange={handleSource}
                        uploadedLabel="File uploaded successfully!"
                    />
                    <p>{sourceFile ? `File name: ${sourceFile.name}` : "No files uploaded yet"}</p>
                </div>
                <div>
                    <label>Target document:</label>
                    <FileUploader
                        types={fileTypes}
                        multiple={true}
                        name="targetFiles"
                        handleChange={handleTarget}
                        uploadedLabel="File(s) uploaded successfully!"
                    />
                    <p>
                        {targetFiles.length > 0 ? (
                            targetFiles.map((file, index) => (
                                <span key={index}>
                                    File name: {file.name}
                                    <br />
                                </span>
                            ))
                        ) : (
                            "No files uploaded yet"
                        )}
                    </p>
                </div>
                {!downloadLink && (
                    <button type="submit">Upload and Format</button>
                )}
            </form>
            <p>{message}</p>
            {downloadLink && (
                <a href={downloadLink} download="formatted_files.zip">
                    <button>Download Formatted Document</button>
                </a>
            )}
        </div>
    )
}

export default Form;