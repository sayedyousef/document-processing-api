<template>
  <div class="min-h-screen bg-gray-50">
    <div class="container mx-auto py-8">
      <h1 class="text-3xl font-bold text-center mb-8">Document Processing Service</h1>
      
      <!-- Processor Selection -->
      <div class="max-w-xl mx-auto mb-6">
        <label class="block text-sm font-medium text-gray-700 mb-2">Select Processor</label>
        <select v-model="processorType" class="w-full px-3 py-2 border border-gray-300 rounded-md">
          <option value="latex_equations">Extract LaTeX Equations</option>
          <option value="word_complete">Complete Word conversion to HTML</option>
          <option value="scan_verify">Scan & Verify Documents</option>
          <option value="word_to_html">Convert to HTML</option>

        </select>
      </div>
      
      <!-- File Upload -->
      <FileUploader 
        @files-selected="handleFiles" 
        @upload="uploadFiles"
        :files="files"
      />
      
      <!-- Job Status -->
      <JobStatus 
        v-if="jobId" 
        :job-id="jobId"
        @completed="handleCompleted"
      />
      
      <!-- Results -->
      <ResultDownload 
        v-if="results.length" 
        :results="results"
        :job-id="completedJobId"
        @reset="resetProcessor"
      />
    </div>
  </div>
</template>

<script>
import { ref } from 'vue'
import FileUploader from './components/FileUploader.vue'
import JobStatus from './components/JobStatus.vue'
import ResultDownload from './components/ResultDownload.vue'
import axios from 'axios'

export default {
  components: {
    FileUploader,
    JobStatus,
    ResultDownload
  },
  setup() {
    const files = ref([])
    const jobId = ref(null)
    const completedJobId = ref(null)
    const results = ref([])
    const processorType = ref('word_to_html')
    
    const handleFiles = (selectedFiles) => {
      files.value = selectedFiles
    }
    
    const uploadFiles = async () => {
      const formData = new FormData()
      files.value.forEach(file => formData.append('files', file))
      formData.append('processor_type', processorType.value)
      
      try {
        const response = await axios.post('http://localhost:8000/api/process', formData)
        jobId.value = response.data.job_id
        console.log('Job started:', jobId.value)
        files.value = [] // Clear files after upload
      } catch (error) {
        console.error('Upload failed:', error)
        alert('Upload failed. Please check if the backend is running.')
      }
    }
    
    const handleCompleted = (jobResults) => {
      console.log('Job completed, results:', jobResults)
      results.value = jobResults
      completedJobId.value = jobId.value  // Save the completed job ID
      jobId.value = null  // Clear current job ID
    }
    
    const resetProcessor = () => {
      files.value = []
      jobId.value = null
      completedJobId.value = null
      results.value = []
    }
    
    return {
      files,
      jobId,
      completedJobId,
      results,
      processorType,
      handleFiles,
      uploadFiles,
      handleCompleted,
      resetProcessor
    }
  }
}
</script>