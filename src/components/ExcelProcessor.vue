<template>
  <div class="excel-processor">
    <h1>Excel Migration Tool</h1>
    
    <div class="upload-section">
      <input type="file" multiple @change="handleFileUpload" accept=".xlsx" />
      <p v-if="files.length > 0">{{ files.length }} files selected</p>
    </div>

    <div v-if="files.length > 0" class="actions">
      <button @click="processFiles" :disabled="processing">
        {{ processing ? 'Processing...' : 'Process Files' }}
      </button>
    </div>

    <div v-if="logs.length > 0" class="logs">
      <h3>Logs:</h3>
      <ul>
        <li v-for="(log, index) in logs" :key="index">{{ log }}</li>
      </ul>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { processNarahuvannya, processVidomist, processBorzhnyky, type ProcessedFile } from '../processor';
import { Constants } from '../constants';

const files = ref<File[]>([]);
const processing = ref(false);
const logs = ref<string[]>([]);

const handleFileUpload = (event: Event) => {
  const target = event.target as HTMLInputElement;
  if (target.files) {
    files.value = Array.from(target.files);
    logs.value = [];
  }
};

const processFiles = async () => {
  processing.value = true;
  logs.value = ['Starting processing...'];
  
  try {
    const zip = new JSZip();
    const tasks = groupFilesByYear(files.value);

    for (const task of tasks) {
      logs.value.push(`Processing task: ${task.sourceYear} -> ${task.targetYear}`);
      const folder = zip.folder(task.targetYear);
      
      for (const file of task.files) {
        let processed: ProcessedFile | null = null;
        
        if (file.name.includes(Constants.PrefixNarahuvannya)) {
          logs.value.push(`Processing ${file.name}...`);
          processed = await processNarahuvannya(file, task.sourceYear, task.targetYear);
        } else if (file.name.includes(Constants.PrefixVidomist)) {
          logs.value.push(`Processing ${file.name}...`);
          processed = await processVidomist(file, task.sourceYear, task.targetYear);
        } else if (file.name.includes(Constants.PrefixBorzhnyky)) {
          logs.value.push(`Processing ${file.name}...`);
          processed = await processBorzhnyky(file, task.sourceYear, task.targetYear);
        }

        if (processed && folder) {
          folder.file(processed.name, processed.buffer);
          logs.value.push(`Finished ${file.name} -> ${processed.name}`);
        }
      }
    }

    const content = await zip.generateAsync({ type: 'blob' });
    saveAs(content, 'migrated_files.zip');
    logs.value.push('All done! Downloading ZIP...');

  } catch (error) {
    console.error(error);
    logs.value.push(`Error: ${error}`);
  } finally {
    processing.value = false;
  }
};

interface Task {
  sourceYear: string;
  targetYear: string;
  files: File[];
}

const groupFilesByYear = (files: File[]): Task[] => {
  const groups: Record<string, File[]> = {};
  const re = /(\d{2})\.xlsx$/;

  for (const file of files) {
    const match = file.name.match(re);
    if (match && match[1]) {
      const suffix = match[1];
      if (!groups[suffix]) groups[suffix] = [];
      groups[suffix].push(file);
    }
  }

  const tasks: Task[] = [];
  for (const suffix in groups) {
    const srcYearInt = parseInt(suffix);
    const targetYearInt = srcYearInt + 1;
    const groupFiles = groups[suffix];
    if (groupFiles) {
        tasks.push({
        sourceYear: `20${suffix}`,
        targetYear: `20${targetYearInt}`,
        files: groupFiles
        });
    }
  }
  return tasks;
};
</script>

<style scoped>
.excel-processor {
  max-width: 600px;
  margin: 0 auto;
  padding: 20px;
  font-family: sans-serif;
}
.upload-section {
  border: 2px dashed #ccc;
  padding: 20px;
  text-align: center;
  margin-bottom: 20px;
}
.actions {
  text-align: center;
  margin-bottom: 20px;
}
button {
  padding: 10px 20px;
  font-size: 16px;
  cursor: pointer;
}
.logs {
  background: #f5f5f5;
  padding: 10px;
  border-radius: 4px;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  margin-bottom: 5px;
}
</style>
