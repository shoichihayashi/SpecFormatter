import React, { useState } from 'react'
import axios from 'axios'

const Form = () => {
    const [sourceFile, setSource] = useState(null)
    const [targetFile, setTarget] = useState(null)
    const [message, setMessage] = useState('')

    const handleSource = (e) => {
        setSource(e.target.files[0])
    }

    const handleTarget = (e) => {
        setTarget(e.target.files[0])
    }

    const handleSubmit = async (e) => {
        e.preventDefault()

        const formData = new FormData()
        formData.append('source_file', sourceFile)
        formData.append('target_file', targetFile)

        try {
            const response = await axios.post('http://localhost:5000/process', formData, {
                headers: {
                    'Content-Type': 'multipart/form-data',
                },
            })
            setMessage(response.data.message)
        } catch(error) {
            console.error('Error uploading files:', error)
            setMessage('Error processing the files')
        }
    }

    return (
        <div>
            <h1>SpecFormatter</h1>
            <form onSubmit={handleSubmit}>
                <div>
                    <label>Source document:</label>
                    <input type="file" onChange={handleSource} />
                </div>
                <div>
                    <label>Target document:</label>
                    <input type="file" onChange={handleTarget} />
                </div>
                <button type="submit">Upload and Format</button>
            </form>
            <p>{message}</p>
        </div>
    )
}

export default Form;