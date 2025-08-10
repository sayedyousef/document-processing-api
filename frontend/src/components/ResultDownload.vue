<template>
  <div class="max-w-xl mx-auto mt-6 p-6 bg-white rounded-lg shadow-md">
    <h3 class="text-lg font-semibold mb-4">Results Ready</h3>
    
    <div class="space-y-3">
      <div 
        v-for="(result, index) in results" 
        :key="index"
        class="flex justify-between items-center p-3 bg-gray-50 rounded-lg"
      >
        <div class="flex items-center space-x-3">
          <svg class="h-5 w-5 text-green-500" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd" />
          </svg>
          <span class="text-sm font-medium">{{ result.filename }}</span>
        </div>
        
        <div class="flex items-center space-x-2">
          <span v-if="result.error" class="text-xs text-red-500">{{ result.error }}</span>
          <button 
            v-else
            @click="downloadFile(index)"
            class="px-3 py-1 bg-blue-500 text-white text-sm rounded hover:bg-blue-600 transition-colors"
          >
            Download
          </button>
        </div>
      </div>
    </div>
    
    <button 
      @click="downloadAll"
      v-if="results.length > 1 && !hasErrors"
      class="mt-4 w-full px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 transition-colors font-semibold"
    >
      Download All as ZIP
    </button>
    
    <button 
      @click="$emit('reset')"
      class="mt-3 w-full px-4 py-2 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 transition-colors"
    >
      Process More Documents
    </button>
  </div>
</template>

<script>
import { computed } from 'vue'

export default {
  props: ['results', 'jobId'],
  emits: ['reset'],
  setup(props) {
    const hasErrors = computed(() => {
      return props.results.some(r => r.error)
    })
    
    const downloadFile = (index) => {
      // Use the jobId from props for download
      const url = `http://localhost:8000/api/download/${props.jobId}/${index}`
      console.log('Downloading from:', url)
      
      // Create download link
      const link = document.createElement('a')
      link.href = url
      link.download = props.results[index].output_filename || props.results[index].filename
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    }
    
    const downloadAll = () => {
      // Download ZIP with all results
      const url = `http://localhost:8000/api/download/${props.jobId}`
      console.log('Downloading ZIP from:', url)
      
      const link = document.createElement('a')
      link.href = url
      link.download = `results_${props.jobId}.zip`
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    }
    
    return {
      downloadFile,
      downloadAll,
      hasErrors
    }
  }
}
</script>