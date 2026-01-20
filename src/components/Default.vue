<script setup lang="ts">
    import { ref } from 'vue'
    import Button from './Button.vue'
    import {useFileDialog} from '@vueuse/core'
    import * as XLSX from 'xlsx'
    
    interface ExcelRow {
        [key: string]: string | number | boolean | null
    }
    const data = ref<ExcelRow[]>([])
    const loading = ref<boolean>(false)
    const error = ref<string | null>(null)
    const fileName = ref<string>('')
    

    const {files, open, reset, onChange} = useFileDialog({
        accept: '.xlsx, .xls',
        multiple: false,
    })

    onChange((files: FileList | null) =>{
        if (!files || files.length === 0) return
        const file: File | null = files.item(0)
        if (!file) return
        importExcel(file)
    })

    const importExcel = async (file: File): Promise<void> => {
        fileName.value = file.name
        loading.value =true
        error.value = null
        data.value = []

        try {
            const arrayBuffer = await file.arrayBuffer()
            const workbook = XLSX.read(arrayBuffer,{type: 'array'})

            const sheetNames = workbook.SheetNames
            const sheetName = sheetNames[0]!
            const worksheet = workbook.Sheets[sheetName]!

            const jsonData = XLSX.utils.sheet_to_json<ExcelRow>(worksheet)
            data.value = jsonData
            console.log("Imported data: ", jsonData)
        } catch (err) {
            error.value = err instanceof Error ? err.message : 'Not Known error'
            console.error('Import Error', err)
        } finally {
            loading.value = false
        }
    }

</script>

<template>
    <div class="section">
        <Button @click="() => open()">Download</Button>
        <Button v-if="files" @click="reset">Reset</Button>
        <div v-if="loading">Downloading</div>
        <div v-if="error" class="error"> Error: {{ error }}</div>
        <template v-if="files">
            <ul>
                <li v-for="value in files">
                    {{ value.name }}
                </li>
            </ul>
        </template>
        <template v-if="data.length">
            <h3>Imported data from {{ fileName }}</h3>
            <table>
                ..in progress
            </table>
        </template>
    </div>
  
</template>

<style scoped>
    .error {
  color: red;
  margin: 10px 0;
}

table {
  border-collapse: collapse;
  margin-top: 20px;
  width: 100%;
}

th, td {
  border: 1px solid #ddd;
  padding: 8px;
  text-align: left;
}

th {
  background-color: #4CAF50;
  color: white;
}
</style>