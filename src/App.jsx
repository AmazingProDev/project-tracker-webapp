import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, AlertCircle, CheckCircle, Clock, Search, Settings, Link as LinkIcon, RotateCw, AlertTriangle, Mail, Trash2 } from 'lucide-react';

function App() {
  const [tasks, setTasks] = useState(() => {
    const savedTasks = localStorage.getItem('tasks');
    return savedTasks ? JSON.parse(savedTasks) : [];
  });
  const [fileName, setFileName] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [filter, setFilter] = useState('all'); // 'all', 'delayed', 'completed'
  const [searchTerm, setSearchTerm] = useState('');
  const [taskNameFilter, setTaskNameFilter] = useState('all'); // New filter for task names
  const [sheetUrl, setSheetUrl] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [showAddTaskModal, setShowAddTaskModal] = useState(false);
  const [showExportModal, setShowExportModal] = useState(false);
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [airtableConfig, setAirtableConfig] = useState({
    apiKey: localStorage.getItem('at_apiKey') || '',
    baseId: localStorage.getItem('at_baseId') || '',
    tableName: localStorage.getItem('at_tableName') || 'Tasks'
  });
  const [exportFilters, setExportFilters] = useState({
    taskName: 'all',
    delayedOnly: false,
    scope: 'all', // 'all' or 'filtered'
    format: 'excel' // 'excel' or 'csv'
  });
  const [newTask, setNewTask] = useState({
    name: '',
    description: '',
    assignee: '',
    status: 'In Progress',
    startDate: '',
    deadline: ''
  });
  const [editingTask, setEditingTask] = useState(null);
  const [customFields, setCustomFields] = useState(() => {
    const saved = localStorage.getItem('customFields');
    return saved ? JSON.parse(saved) : [];
  });
  const [columns, setColumns] = useState(() => {
    const saved = localStorage.getItem('tableColumns');
    return saved ? JSON.parse(saved) : ['Task Name', 'Assignee', 'Status', 'Start Date', 'Deadline'];
  });
  const [hiddenColumns, setHiddenColumns] = useState(() => {
    const saved = localStorage.getItem('hiddenColumns');
    return saved ? JSON.parse(saved) : [];
  });
  const [showColumnManager, setShowColumnManager] = useState(false);
  const [newColumnName, setNewColumnName] = useState('');
  const [newColumnType, setNewColumnType] = useState('text');
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });
  const tableRef = React.useRef(null);

  const scrollToTable = () => {
    if (tableRef.current) {
      tableRef.current.scrollIntoView({ behavior: 'smooth' });
    }
  };

  // Save tasks to localStorage whenever they change
  React.useEffect(() => {
    localStorage.setItem('tasks', JSON.stringify(tasks));
  }, [tasks]);

  // Handle File Upload
  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) processFile(file);
  };

  const processFile = (file) => {
    setFileName(file.name);
    setError(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const arrayBuffer = evt.target.result;
      // Force UTF-8 (65001) to ensure CSVs are read correctly
      // Enable cellDates to let XLSX parse dates correctly
      const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, codepage: 65001 });
      processWorkbook(wb);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleSheetLoad = async () => {
    if (!sheetUrl) return;
    setIsLoading(true);
    setError(null);

    try {
      // Convert standard Google Sheet URL to CSV export URL
      // From: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit...
      // To:   https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/export?format=csv

      let exportUrl = sheetUrl;

      // Handle "Published to the web" URLs (e.g., .../pubhtml)
      if (sheetUrl.includes('/pubhtml')) {
        exportUrl = sheetUrl.replace(/\/pubhtml.*/, '/pub?output=csv');
      }
      // Handle standard "Edit" URLs (e.g., .../edit)
      else {
        const match = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
        // Ensure we don't capture 'e' from /d/e/... URLs if the user pasted a partial link
        if (match && match[1] && match[1] !== 'e') {
          exportUrl = `https://docs.google.com/spreadsheets/d/${match[1]}/export?format=csv`;
        }
      }

      const response = await fetch(exportUrl);
      if (!response.ok) throw new Error('Failed to fetch Google Sheet. Make sure it is public.');

      const arrayBuffer = await response.arrayBuffer();
      // Force UTF-8 (65001) here as well
      const wb = XLSX.read(arrayBuffer, { type: 'array', codepage: 65001 });
      processWorkbook(wb);
      setFileName('Google Sheet Data');
    } catch (err) {
      console.error(err);
      setError('Could not load Google Sheet. Ensure "Anyone with the link" is set to Viewer.');
    } finally {
      setIsLoading(false);
    }
  };

  const calculateTimeMetrics = (status, startDate, deadline) => {
    // Rule 1: If status = Done then Delay / running days = OK
    const isCompleted = ['Done', 'Completed', 'Terminé', 'Finished'].includes(status);
    if (isCompleted) {
      return { display: 'OK', className: 'badge badge-green', raw: 0 };
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Helper to check if two dates are the same day
    const isSameDay = (d1, d2) => {
      return d1.getFullYear() === d2.getFullYear() &&
        d1.getMonth() === d2.getMonth() &&
        d1.getDate() === d2.getDate();
    };

    // Rule 2: If status = In progress and deadline is not empty
    if (deadline) {
      if (isSameDay(today, deadline)) {
        // Deadline is Today
        return { display: 'Deadline Now', className: 'delay-cell', raw: 0 };
      }

      const diffTime = today - deadline;
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      if (diffDays > 0) {
        // Deadline passed
        return { display: `Delayed (+${diffDays} days)`, className: 'delay-cell', raw: diffDays };
      } else {
        // Deadline in future
        return { display: `${Math.abs(diffDays)} days left`, className: 'badge badge-blue', raw: diffDays };
      }
    }

    // Rule 3: If status = In progress and deadline is empty
    if (startDate) {
      const diffStart = today - startDate;
      const runningDays = Math.ceil(diffStart / (1000 * 60 * 60 * 24));

      // New Rule: If running days > 15 and deadline is empty then colorize Delay / running days by red
      if (runningDays > 15) {
        return { display: `Running (${runningDays} days)`, className: 'delay-cell', raw: runningDays };
      }

      return { display: `Running (${runningDays} days)`, className: 'badge badge-green', raw: runningDays };
    }

    return { display: '-', className: '', raw: 0 };
  };

  const handleAirtableLoad = async () => {
    if (!airtableConfig.apiKey || !airtableConfig.baseId) {
      setError('Please configure Airtable Settings first (click the gear icon).');
      return;
    }

    setIsLoading(true);
    setError(null);

    try {
      const response = await fetch(`https://api.airtable.com/v0/${airtableConfig.baseId}/${airtableConfig.tableName}`, {
        headers: {
          'Authorization': `Bearer ${airtableConfig.apiKey}`
        }
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        console.error('Airtable Error Details:', errorData);

        const errorMessage = errorData.error?.message || errorData.error || response.statusText;

        if (response.status === 401) throw new Error('Unauthorized: Check your API Key (should start with "pat...")');
        if (response.status === 403) throw new Error(`Permission Denied (403): ${errorMessage}. Check Token Scopes (data.records:read) and Base Access.`);
        if (response.status === 404) throw new Error('Not Found: Check your Base ID (starts with "app...") and Table Name');

        throw new Error(`Airtable Error (${response.status}): ${errorMessage}`);
      }

      const data = await response.json();

      // Extract dynamic columns from the first record (or all records to be safe)
      if (data.records.length > 0) {
        // Get all unique keys from all records fields
        const allKeys = new Set();
        data.records.forEach(record => {
          Object.keys(record.fields).forEach(key => allKeys.add(key));
        });

        // Prioritize standard columns order if they exist
        const standardOrder = ['Task Name', 'Description', 'Assignee', 'Status', 'Start Date', 'Deadline'];
        const sortedKeys = Array.from(allKeys).sort((a, b) => {
          const indexA = standardOrder.indexOf(a);
          const indexB = standardOrder.indexOf(b);
          if (indexA !== -1 && indexB !== -1) return indexA - indexB;
          if (indexA !== -1) return -1;
          if (indexB !== -1) return 1;
          return a.localeCompare(b);
        });

        // Merge with existing columns to preserve order
        setColumns(prevColumns => {
          const newColumns = [...prevColumns];
          sortedKeys.forEach(key => {
            if (!newColumns.includes(key)) {
              newColumns.push(key);
            }
          });
          // Optional: Remove columns that no longer exist? 
          // For now, let's keep them to avoid losing preferences if a column is temporarily missing
          return newColumns;
        });
      }

      const formattedTasks = data.records.map((record, index) => {
        const fields = record.fields;

        // Helper to parse date string YYYY-MM-DD to Date object
        const parseDateString = (dateStr) => {
          if (!dateStr) return null;
          return new Date(dateStr);
        };

        const startDate = parseDateString(fields['Start Date']);
        const deadlineDate = parseDateString(fields['Deadline']);

        // Use centralized time metrics calculation
        const timeMetrics = calculateTimeMetrics(
          fields['Status'] || 'Pending',
          startDate,
          deadlineDate
        );

        // Format date for display
        const formatDateDisplay = (date) => {
          if (!date) return '-';
          const day = String(date.getDate()).padStart(2, '0');
          const month = String(date.getMonth() + 1).padStart(2, '0');
          const year = date.getFullYear();
          return `${month}/${day}/${year}`;
        };

        return {
          id: record.id, // Use Airtable ID
          ...fields, // Spread all original fields
          name: fields['Task Name'] || 'Untitled', // Keep internal mapping for logic
          description: fields['Description'] || '',
          assignee: fields['Assignee'] || 'Unassigned',
          status: fields['Status'] || 'Pending',
          startDate: formatDateDisplay(startDate),
          deadline: formatDateDisplay(deadlineDate),
          timeMetrics: timeMetrics
        };
      });

      setTasks(formattedTasks);
      setFileName(`Airtable: ${airtableConfig.tableName}`);
    } catch (err) {
      console.error(err);
      setError(err.message);
    } finally {
      setIsLoading(false);
    }
  };

  const handleAddTask = () => {
    if (!newTask.name || !newTask.assignee || !newTask.startDate) {
      alert('Please fill in at least Task Name, Assignee, and Start Date');
      return;
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const parseLocal = (dateStr) => {
      if (!dateStr) return null;
      const [y, m, d] = dateStr.split('-').map(Number);
      return new Date(y, m - 1, d);
    };

    const startDate = parseLocal(newTask.startDate);
    const deadlineDate = parseLocal(newTask.deadline);

    let delayDays = 0;
    let isDelayed = false;
    let runningDays = 0;

    if (startDate) {
      const diffStart = Math.abs(today - startDate);
      runningDays = Math.ceil(diffStart / (1000 * 60 * 60 * 24));
    }

    // Only calculate delays for non-completed tasks
    const isCompleted = ['Done', 'Completed', 'Terminé', 'Finished'].includes(newTask.status);

    if (!isCompleted && deadlineDate) {
      if (today > deadlineDate) {
        const diffTime = Math.abs(today - deadlineDate);
        delayDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        isDelayed = true;
      }
    }

    // Format dates manually for display (DD/MM/YYYY)
    const formatDateManual = (date) => {
      if (!date) return '-';
      const day = String(date.getDate()).padStart(2, '0');
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const year = date.getFullYear();
      return `${day}/${month}/${year}`;
    };

    // Save to Airtable if configured
    if (airtableConfig.apiKey && airtableConfig.baseId) {
      const airtableData = {
        records: [
          {
            fields: {
              "Task Name": newTask.name,
              "Description": newTask.description,
              "Assignee": newTask.assignee,
              "Status": newTask.status,
              "Start Date": startDate ? startDate.toISOString().split('T')[0] : null,
              "Deadline": deadlineDate ? deadlineDate.toISOString().split('T')[0] : null,
              // Add custom fields
              ...customFields.reduce((acc, field, index) => {
                const value = newTask[`custom_${index}`];
                if (value) {
                  acc[field.name] = field.type === 'number' ? Number(value) : value;
                }
                return acc;
              }, {})
            }
          }
        ]
      };

      fetch(`https://api.airtable.com/v0/${airtableConfig.baseId}/${airtableConfig.tableName}`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${airtableConfig.apiKey}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(airtableData)
      })
        .then(async response => {
          if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            console.error('Airtable Save Error:', errorData);
            const errorMessage = errorData.error?.message || errorData.error || response.statusText;

            if (response.status === 401) throw new Error('Unauthorized: Check API Key');
            if (response.status === 403) throw new Error(`Permission Denied (403): ${errorMessage}. Check "data.records:write" scope.`);
            if (response.status === 404) throw new Error('Not Found: Check Base ID/Table Name');
            if (response.status === 422) throw new Error(`Invalid Data (422): ${errorMessage}. Check field names match exactly.`);

            throw new Error(`Error (${response.status}): ${errorMessage}`);
          }
          return response.json();
        })
        .then(data => {
          console.log('Saved to Airtable:', data);
          alert('Task saved to Airtable successfully!');
        })
        .catch(err => {
          console.error('Error saving to Airtable:', err);
          alert(`Failed to save: ${err.message}`);
        });
    }

    const task = {
      id: tasks.length,
      name: newTask.name,
      description: newTask.description,
      assignee: newTask.assignee,
      status: newTask.status,
      startDate: formatDateManual(startDate),
      deadline: deadlineDate ? formatDateManual(deadlineDate) : '-',
      timeMetrics: calculateTimeMetrics(newTask.status, startDate, deadlineDate),
      // Include custom fields
      ...customFields.reduce((acc, field, index) => {
        acc[`custom_${index}`] = newTask[`custom_${index}`] || '';
        return acc;
      }, {})
    };

    setTasks([...tasks, task]);
    setShowAddTaskModal(false);
    setNewTask({
      name: '',
      description: '',
      assignee: '',
      status: 'In Progress',
      startDate: '',
      deadline: ''
    });
  };

  const handleEditClick = (task) => {
    setEditingTask(task);
    // Convert DD/MM/YYYY to YYYY-MM-DD for input fields
    const toInputDate = (dateStr) => {
      if (!dateStr || dateStr === '-') return '';
      const [day, month, year] = dateStr.split('/');
      return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
    };

    setNewTask({
      name: task.name,
      description: task.description || '',
      assignee: task.assignee,
      status: task.status,
      startDate: toInputDate(task.startDate),
      deadline: toInputDate(task.deadline),
      // Include custom fields
      ...customFields.reduce((acc, field, index) => {
        const value = task[`custom_${index}`];
        // Convert date format for date inputs if needed
        if (field.type === 'date' && value && value !== '-') {
          acc[`custom_${index}`] = toInputDate(value);
        } else {
          acc[`custom_${index}`] = value || '';
        }
        return acc;
      }, {})
    });
    setShowAddTaskModal(true);
  };

  const handleDeleteTask = async () => {
    if (!editingTask || !confirm('Are you sure you want to delete this task?')) return;

    if (airtableConfig.apiKey && airtableConfig.baseId && editingTask.id) {
      try {
        const response = await fetch(`https://api.airtable.com/v0/${airtableConfig.baseId}/${airtableConfig.tableName}/${editingTask.id}`, {
          method: 'DELETE',
          headers: {
            'Authorization': `Bearer ${airtableConfig.apiKey}`
          }
        });

        if (!response.ok) throw new Error('Failed to delete from Airtable');
        console.log('Deleted from Airtable');
      } catch (err) {
        console.error('Error deleting from Airtable:', err);
        alert('Failed to delete from Airtable, but removing locally.');
      }
    }

    setTasks(tasks.filter(t => t.id !== editingTask.id));
    setShowAddTaskModal(false);
    setEditingTask(null);
    setNewTask({ name: '', description: '', assignee: '', status: 'In Progress', startDate: '', deadline: '' });
  };

  const handleUpdateTask = async () => {
    if (!newTask.name || !newTask.assignee || !newTask.startDate) {
      alert('Please fill in at least Task Name, Assignee, and Start Date');
      return;
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const parseLocal = (dateStr) => {
      if (!dateStr) return null;
      const [y, m, d] = dateStr.split('-').map(Number);
      return new Date(y, m - 1, d);
    };

    const startDate = parseLocal(newTask.startDate);
    const deadlineDate = parseLocal(newTask.deadline);

    let delayDays = 0;
    let isDelayed = false;
    let runningDays = 0;

    if (startDate) {
      const diffStart = Math.abs(today - startDate);
      runningDays = Math.ceil(diffStart / (1000 * 60 * 60 * 24));
    }

    const isCompleted = ['Done', 'Completed', 'Terminé', 'Finished'].includes(newTask.status);

    if (!isCompleted && deadlineDate) {
      if (today > deadlineDate) {
        const diffTime = Math.abs(today - deadlineDate);
        delayDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        isDelayed = true;
      }
    }

    const formatDateManual = (date) => {
      if (!date) return '-';
      const day = String(date.getDate()).padStart(2, '0');
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const year = date.getFullYear();
      return `${day}/${month}/${year}`;
    };

    // Update in Airtable
    if (airtableConfig.apiKey && airtableConfig.baseId && editingTask.id) {
      const airtableData = {
        fields: {
          "Task Name": newTask.name,
          "Description": newTask.description,
          "Assignee": newTask.assignee,
          "Status": newTask.status,
          "Start Date": startDate ? startDate.toISOString().split('T')[0] : null,
          "Deadline": deadlineDate ? deadlineDate.toISOString().split('T')[0] : null,
          // Add custom fields
          ...customFields.reduce((acc, field, index) => {
            const value = newTask[`custom_${index}`];
            if (value) {
              acc[field.name] = field.type === 'number' ? Number(value) : value;
            }
            return acc;
          }, {})
        }
      };

      try {
        const response = await fetch(`https://api.airtable.com/v0/${airtableConfig.baseId}/${airtableConfig.tableName}/${editingTask.id}`, {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${airtableConfig.apiKey}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(airtableData)
        });

        if (!response.ok) throw new Error('Failed to update Airtable');
        console.log('Updated Airtable');
      } catch (err) {
        console.error('Error updating Airtable:', err);
        alert('Failed to update Airtable, but updating locally.');
      }
    }

    const updatedTask = {
      ...editingTask,
      name: newTask.name,
      description: newTask.description,
      assignee: newTask.assignee,
      status: newTask.status,
      startDate: formatDateManual(startDate),
      deadline: deadlineDate ? formatDateManual(deadlineDate) : '-',
      timeMetrics: calculateTimeMetrics(newTask.status, startDate, deadlineDate),
      // Include updated custom fields
      ...customFields.reduce((acc, field, index) => {
        acc[`custom_${index}`] = newTask[`custom_${index}`] || '';
        return acc;
      }, {})
    };

    setTasks(tasks.map(t => t.id === editingTask.id ? updatedTask : t));
    setShowAddTaskModal(false);
    setEditingTask(null);
    setNewTask({ name: '', description: '', assignee: '', status: 'In Progress', startDate: '', deadline: '' });
  };

  const handleExportToExcel = async () => {
    let tasksToExport = [];

    // Determine source based on scope
    if (exportFilters.scope === 'filtered') {
      tasksToExport = filteredTasks;
    } else {
      tasksToExport = tasks;

      // Apply manual filters if 'all' scope but specific filters selected (legacy support or specific overrides)
      if (exportFilters.taskName !== 'all') {
        tasksToExport = tasksToExport.filter(t => t.name === exportFilters.taskName);
      }
      if (exportFilters.delayedOnly) {
        tasksToExport = tasksToExport.filter(t => t.isDelayed);
      }
    }

    if (tasksToExport.length === 0) {
      alert('No tasks to export with the selected options');
      return;
    }

    // Prepare data for export
    const exportData = tasksToExport.map(task => ({
      'Task Name': task.name,
      'Description': task.description || '-',
      'Assignee': task.assignee,
      'Status': task.status,
      'Start Date': task.startDate,
      'Deadline': task.deadline,
      'Delay (Days)': task.isDelayed ? task.delay : '-',
      'Running (Days)': task.runningDays || '-',
      // Include custom fields
      ...Object.keys(task).filter(k => k.startsWith('custom_')).reduce((acc, key) => {
        // Find column name for this custom field
        const colIndex = parseInt(key.split('_')[1]);
        const colName = customFields[colIndex]?.name || key;
        acc[colName] = task[key];
        return acc;
      }, {})
    }));

    // Create workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Tasks');

    // Generate filename
    const timestamp = new Date().toISOString().split('T')[0];
    let filename = `tasks_export_${timestamp}`;

    if (exportFilters.scope === 'filtered') {
      filename += '_filtered';
    } else {
      if (exportFilters.taskName !== 'all') {
        filename = `${exportFilters.taskName.replace(/\s+/g, '_')}_${timestamp}`;
      }
      if (exportFilters.delayedOnly) {
        filename += '_delayed';
      }
    }

    try {
      // Determine file extension and MIME type
      let extension, mimeType, bookType;
      if (exportFilters.format === 'csv') {
        extension = '.csv';
        mimeType = 'text/csv';
        bookType = 'csv';
      } else {
        // Excel and Numbers both use .xlsx
        extension = '.xlsx';
        mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        bookType = 'xlsx';
      }

      // Check if File System Access API is supported
      if ('showSaveFilePicker' in window) {
        try {
          // Show Save As dialog
          const fileHandle = await window.showSaveFilePicker({
            suggestedName: filename + extension,
            types: [{
              description: exportFilters.format === 'csv' ? 'CSV File' : 'Excel File',
              accept: { [mimeType]: [extension] }
            }]
          });

          // Generate file blob
          const wbout = XLSX.write(wb, { bookType: bookType, type: 'array' });
          const blob = new Blob([wbout], { type: mimeType });

          // Write to file
          const writable = await fileHandle.createWritable();
          await writable.write(blob);
          await writable.close();
        } catch (err) {
          // User cancelled or error occurred, fallback to standard download
          if (err.name !== 'AbortError') {
            console.error('Save dialog error:', err);
            // Fallback to standard download
            XLSX.writeFile(wb, filename + extension, { bookType: bookType });
          }
        }
      } else {
        // Fallback for browsers that don't support File System Access API
        XLSX.writeFile(wb, filename + extension, { bookType: bookType });
      }
    } catch (err) {
      console.error('Export failed:', err);
      alert('Failed to export data. Please try again.');
    }
  };

  const processWorkbook = (wb) => {
    const wsname = wb.SheetNames[0];
    const ws = wb.Sheets[wsname];
    // Use cellDates: true to get JS Date objects directly
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, cellDates: true, dateNF: 'dd/mm/yyyy' });

    // Helper to normalize strings (remove accents, lowercase)
    const normalizeStr = (str) => {
      return String(str || '')
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim();
    };

    // Smart Header Detection: Scan first 20 rows to find the header
    let headerRowIndex = -1;
    let headerRow = [];

    for (let i = 0; i < Math.min(data.length, 20); i++) {
      const row = data[i].map(normalizeStr);
      // Check if this row looks like a header row
      // It should contain at least 'task'/'nom' and 'status'/'statut'
      const hasName = row.some(h => h.includes('task') || h.includes('nom') || h.includes('tache'));
      const hasStatus = row.some(h => h.includes('status') || h.includes('statut') || h.includes('etat'));

      if (hasName && hasStatus) {
        headerRowIndex = i;
        headerRow = data[i]; // Use original case for display
        break;
      }
    }

    // If no header found, fallback to first row
    if (headerRowIndex === -1) {
      headerRowIndex = 0;
      headerRow = data[0];
    }

    // Update columns state with found headers
    // Filter out empty headers
    const validHeaders = headerRow.filter(h => h && String(h).trim() !== '');
    setColumns(validHeaders);
    localStorage.setItem('tableColumns', JSON.stringify(validHeaders));

    // Find special column indices for stats calculation
    const normalizedHeaders = headerRow.map(normalizeStr);
    const nameIdx = normalizedHeaders.findIndex(h => h.includes('task') || h.includes('nom') || h.includes('tache'));
    const assigneeIdx = normalizedHeaders.findIndex(h => h.includes('assign') || h.includes('responsable'));
    const statusIdx = normalizedHeaders.findIndex(h => h.includes('status') || h.includes('statut') || h.includes('etat'));
    const startIdx = normalizedHeaders.findIndex(h => h.includes('start') || h.includes('debut'));
    const deadlineIdx = normalizedHeaders.findIndex(h => h.includes('deadline') || h.includes('fin') || h.includes('echeance'));

    const IDX = {
      name: nameIdx !== -1 ? nameIdx : -1,
      assignee: assigneeIdx !== -1 ? assigneeIdx : -1,
      status: statusIdx !== -1 ? statusIdx : -1,
      start: startIdx !== -1 ? startIdx : -1,
      deadline: deadlineIdx !== -1 ? deadlineIdx : -1
    };

    // Data starts after the header row
    const rows = data.slice(headerRowIndex + 1);

    const formattedTasks = rows.map((row, index) => {
      const getVal = (idx) => row[idx] || '';

      // Create dynamic task object based on headers
      const taskObj = { id: index };
      headerRow.forEach((header, colIdx) => {
        if (header && String(header).trim() !== '') {
          // Handle dates if this column is a date column
          let val = row[colIdx];

          // If we suspect it's a date column (based on header name or value type)
          // We can try to parse it, but since we use cellDates: true, it should be a Date object if valid
          if (val instanceof Date) {
            val = formatDate(val);
          } else if (colIdx === IDX.start || colIdx === IDX.deadline) {
            const parsed = parseDate(val);
            const formatted = formatDate(parsed);
            if (colIdx === IDX.deadline) {
              console.log(`Row ${index} Deadline:`, { raw: val, parsed, formatted });
            }
            val = formatted;
          }

          taskObj[header] = val || '';
        }
      });

      // Extract special fields for stats
      const rawStartDate = IDX.start !== -1 ? getVal(IDX.start) : null;
      const startDate = parseDate(rawStartDate);
      const rawDeadline = IDX.deadline !== -1 ? getVal(IDX.deadline) : null;
      const deadlineDate = parseDate(rawDeadline);
      const taskStatus = IDX.status !== -1 ? getVal(IDX.status) : '';
      const taskName = IDX.name !== -1 ? getVal(IDX.name) : '';

      // Calculate delay in JS
      let delayDays = 0;
      let isDelayed = false;
      let runningDays = 0;

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      if (startDate) {
        const diffStart = Math.abs(today - startDate);
        runningDays = Math.ceil(diffStart / (1000 * 60 * 60 * 24));
      }

      // Check if task is completed
      const normalize = (s) => String(s || '')
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim();

      const isCompleted = ['done', 'completed', 'termine', 'finished'].includes(normalize(taskStatus));

      // Only calculate delays for non-completed tasks
      if (!isCompleted && deadlineDate) {
        // If deadline exists, Delay = Today - Deadline (if passed)
        if (today > deadlineDate) {
          const diffTime = Math.abs(today - deadlineDate);
          delayDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
          isDelayed = true;
        }
      }

      // Add internal fields for app logic (stats, filtering)
      // These won't be displayed in the dynamic table unless they match a header
      const timeMetrics = calculateTimeMetrics(taskStatus, startDate, deadlineDate);
      taskObj._internal = {
        name: taskName,
        status: taskStatus,
        startDate: startDate,
        deadline: deadlineDate,
        delay: delayDays,
        isDelayed: isDelayed,
        runningDays: runningDays,
        timeMetrics: timeMetrics
      };

      // Map standard fields for compatibility with existing components
      taskObj.name = taskName; // Used for filtering
      taskObj.status = taskStatus; // Used for filtering
      taskObj.isDelayed = isDelayed; // Used for filtering
      taskObj.deadline = formatDate(rawDeadline); // Used for display logic sometimes
      taskObj.timeMetrics = timeMetrics; // Expose for stats/table

      return taskObj;
    }).filter(t => t.name); // Filter empty rows (based on identified name column)

    setTasks(formattedTasks);
  };

  const parseDate = (val) => {
    if (!val) return null;

    // 0. Handle JS Date Objects (from cellDates: true)
    if (val instanceof Date) {
      return !isNaN(val.getTime()) ? val : null;
    }

    // 1. Handle Excel/Google Sheets Serial Date (numbers)
    if (typeof val === 'number') {
      // Google Sheets uses the same epoch as Excel (Dec 30, 1899)
      // Serial number represents days since epoch
      // 45667 should be Nov 1, 2025
      const EXCEL_EPOCH = new Date(1899, 11, 30); // Dec 30, 1899
      const date = new Date(EXCEL_EPOCH.getTime() + val * 86400 * 1000);
      return date;
    }

    // 2. Handle String Dates - ALWAYS treat as DD/MM/YYYY
    if (typeof val === 'string') {
      const cleanVal = val.trim();

      // Match DD/MM/YYYY or DD-MM-YYYY or DD.MM.YYYY (with optional time)
      const parts = cleanVal.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);

      if (parts) {
        const day = parseInt(parts[1], 10);
        const month = parseInt(parts[2], 10) - 1; // Months are 0-indexed
        const year = parseInt(parts[3], 10);
        const date = new Date(year, month, day);
        return date;
      }

      // Try standard date parsing as fallback
      const date = new Date(cleanVal);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }

    return null;
  };



  const formatDate = (val) => {
    const date = parseDate(val);
    if (!date) return '-';

    // Format as DD/MM/YYYY
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();

    return `${day}/${month}/${year}`;
  };

  // Derived State
  const filteredTasks = useMemo(() => {
    let result = tasks.filter(task => {
      const matchesSearch = task.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
        task.assignee.toLowerCase().includes(searchTerm.toLowerCase());

      const matchesTaskName = taskNameFilter === 'all' || task.name === taskNameFilter;

      // Normalization helper for status matching
      const normalize = (s) => String(s || '')
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim();

      if (filter === 'delayed') return matchesSearch && matchesTaskName && task.isDelayed;
      if (filter === 'completed') {
        const s = normalize(task.status);
        return matchesSearch && matchesTaskName && (s === 'done' || s === 'completed' || s === 'termine' || s === 'finished');
      }
      if (filter === 'in-progress') {
        const s = normalize(task.status);
        return matchesSearch && matchesTaskName && (s === 'in progress' || s === 'en cours' || s === 'ongoing' || s === 'doing');
      }
      if (filter === 'reappeared') {
        const s = normalize(task.status);
        return matchesSearch && matchesTaskName && (s === 'reappeared' || s === 'réapparu');
      }
      if (filter === 'deadlineNow') {
        return matchesSearch && matchesTaskName && task.timeMetrics?.display === 'Deadline Now';
      }

      return matchesSearch && matchesTaskName;
    });

    // Apply Sorting
    if (sortConfig.key) {
      result.sort((a, b) => {
        let aValue = a[sortConfig.key];
        let bValue = b[sortConfig.key];

        // Handle special cases or nulls
        if (aValue === undefined || aValue === null) aValue = '';
        if (bValue === undefined || bValue === null) bValue = '';

        // Special sort for timeMetrics
        if (sortConfig.key === 'timeMetrics') {
          const aRaw = a.timeMetrics?.raw || -999999;
          const bRaw = b.timeMetrics?.raw || -999999;
          return sortConfig.direction === 'asc' ? aRaw - bRaw : bRaw - aRaw;
        }

        // Numeric sorting
        if (typeof aValue === 'number' && typeof bValue === 'number') {
          return sortConfig.direction === 'asc' ? aValue - bValue : bValue - aValue;
        }

        // String sorting
        const aString = String(aValue).toLowerCase();
        const bString = String(bValue).toLowerCase();

        if (aString < bString) {
          return sortConfig.direction === 'asc' ? -1 : 1;
        }
        if (aString > bString) {
          return sortConfig.direction === 'asc' ? 1 : -1;
        }
        return 0;
      });
    }

    return result;
  }, [tasks, filter, searchTerm, taskNameFilter, sortConfig]);

  const handleSort = (key) => {
    let direction = 'asc';
    if (sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
  };

  // Get unique task names for the filter dropdown
  const uniqueTaskNames = useMemo(() => {
    return ['all', ...new Set(tasks.map(t => t.name).filter(Boolean))];
  }, [tasks]);

  const stats = useMemo(() => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const isSameDay = (d1, d2) => {
      return d1.getFullYear() === d2.getFullYear() &&
        d1.getMonth() === d2.getMonth() &&
        d1.getDate() === d2.getDate();
    };

    const normalize = (s) => String(s || '')
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();

    return {
      total: tasks.length,
      delayed: tasks.filter(t => t.isDelayed).length,
      completed: tasks.filter(t => {
        const s = normalize(t.status);
        return s === 'done' || s === 'completed' || s === 'termine' || s === 'finished';
      }).length,
      deadlineNow: tasks.filter(t => t.timeMetrics?.display === 'Deadline Now').length,
      inProgress: tasks.filter(t => {
        const s = normalize(t.status);
        const isInProgress = s === 'in progress' || s === 'en cours' || s === 'ongoing' || s === 'doing';

        // Exclude if Deadline is Now (as it's counted in deadlineNow)
        if (isInProgress && t.timeMetrics?.display === 'Deadline Now') {
          return false;
        }
        return isInProgress;
      }).length
    };
  }, [tasks]);

  // Task-specific statistics
  const taskTypeStats = useMemo(() => {
    const stats = {};
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const isSameDay = (d1, d2) => {
      return d1.getFullYear() === d2.getFullYear() &&
        d1.getMonth() === d2.getMonth() &&
        d1.getDate() === d2.getDate();
    };

    const normalize = (s) => String(s || '')
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();

    tasks.forEach(task => {
      if (!task.name) return;

      if (!stats[task.name]) {
        stats[task.name] = {
          total: 0,
          delayed: 0,
          completed: 0,
          inProgress: 0,
          reappeared: 0, // Added 'reappeared'
          deadlineNow: 0 // Added 'deadlineNow'
        };
      }

      stats[task.name].total++;

      if (task.isDelayed) {
        stats[task.name].delayed++;
      }

      let isDeadlineNow = false;
      if (task.timeMetrics?.display === 'Deadline Now') {
        stats[task.name].deadlineNow++;
        isDeadlineNow = true;
      }

      const s = normalize(task.status);
      if (s === 'done' || s === 'completed' || s === 'termine' || s === 'finished') {
        stats[task.name].completed++;
      } else if ((s === 'in progress' || s === 'en cours' || s === 'ongoing' || s === 'doing') && !isDeadlineNow) {
        stats[task.name].inProgress++;
      } else if (s === 'reappeared' || s === 'réapparu') { // Added condition for 'reappeared'
        stats[task.name].reappeared++;
      }
    });

    return stats;
  }, [tasks]);

  return (
    <div className="container">
      <header className="header">
        <div>
          <h1 className="title">Project Tracker</h1>
          <p style={{ color: 'var(--color-text-muted)' }}>
            {fileName ? `Viewing: ${fileName}` : 'Created by Abdelilah ELQORCHI Email: e@iam.ma'}
          </p>
        </div>
        {tasks.length > 0 && (
          <div style={{ display: 'flex', gap: '0.5rem' }}>
            <button className="btn btn-primary" onClick={() => setShowAddTaskModal(true)}>
              + Add Task
            </button>
            <button className="btn btn-primary" onClick={() => setShowExportModal(true)}>
              Export Data
            </button>
            <button className="btn btn-outline" onClick={() => setShowSettingsModal(true)} title="Airtable Settings">
              <Settings size={18} />
            </button>
            <button className="btn btn-outline" onClick={() => { setTasks([]); setFileName(null); setSheetUrl(''); }}>
              Load New Source
            </button>
          </div>
        )}
      </header>

      {tasks.length === 0 ? (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          {/* Excel Upload */}
          <div
            className={`upload-zone ${isDragging ? 'drag-active' : ''}`}
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
          >
            <FileSpreadsheet className="upload-icon" />
            <div>
              <h3 style={{ fontSize: '1.25rem', marginBottom: '0.5rem' }}>Drop your Excel file here</h3>
              <p style={{ color: 'var(--color-text-muted)', marginBottom: '1.5rem' }}>or click to browse</p>
              <label className="btn btn-primary">
                Choose File
                <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} hidden />
              </label>
            </div>
          </div>

          {/* Divider */}
          <div className="text-center" style={{ position: 'relative' }}>
            <span style={{ background: 'var(--color-bg)', padding: '0 1rem', color: 'var(--color-text-muted)', position: 'relative', zIndex: 1 }}>OR</span>
            <div style={{ position: 'absolute', top: '50%', left: 0, right: 0, height: '1px', background: 'var(--color-border)', zIndex: 0 }}></div>
          </div>

          {/* Airtable Load */}
          <div className="card" style={{ textAlign: 'center' }}>
            <div style={{ marginBottom: '1rem' }}>
              <div
                onClick={() => setShowSettingsModal(true)}
                style={{
                  width: '48px',
                  height: '48px',
                  background: 'var(--color-primary)',
                  borderRadius: '12px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  margin: '0 auto 1rem',
                  color: 'white',
                  cursor: 'pointer',
                  transition: 'transform 0.2s'
                }}
                onMouseOver={(e) => e.currentTarget.style.transform = 'scale(1.05)'}
                onMouseOut={(e) => e.currentTarget.style.transform = 'scale(1)'}
                title="Open Settings"
              >
                <Settings size={24} />
              </div>
              <h3 style={{ fontSize: '1.25rem', marginTop: '0.5rem' }}>Load from Airtable</h3>
              <p style={{ color: 'var(--color-text-muted)', fontSize: '0.875rem' }}>
                Fetch tasks directly from your connected Airtable Base
              </p>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '1rem' }}>
              {!airtableConfig.apiKey ? (
                <button
                  className="btn btn-outline"
                  onClick={() => setShowSettingsModal(true)}
                >
                  Configure Airtable Settings
                </button>
              ) : (
                <button
                  className="btn btn-primary"
                  onClick={handleAirtableLoad}
                  disabled={isLoading}
                  style={{ minWidth: '200px' }}
                >
                  {isLoading ? 'Loading Tasks...' : 'Load Tasks from Airtable'}
                </button>
              )}
            </div>

            {error && (
              <p style={{ color: 'var(--color-danger)', marginTop: '1rem', fontSize: '0.875rem' }}>
                {error}
              </p>
            )}
          </div>
        </div>
      ) : (
        <>
          {/* Stats Grid */}
          <div className="stats-grid">
            <div
              className="card stat-card area-total"
              onClick={() => { setFilter('all'); setTaskNameFilter('all'); scrollToTable(); }}
              style={{ cursor: 'pointer', transition: 'transform 0.2s', justifyContent: 'center', alignItems: 'center', textAlign: 'center' }}
              onMouseOver={(e) => e.currentTarget.style.transform = 'scale(1.02)'}
              onMouseOut={(e) => e.currentTarget.style.transform = 'scale(1)'}
            >
              <span className="stat-label">Total Tasks</span>
              <span className="stat-value">{stats.total}</span>
            </div>

            <div
              className="card stat-card area-delayed"
              style={{ borderLeft: '4px solid var(--color-danger)', cursor: 'pointer', transition: 'transform 0.2s' }}
              onClick={() => { setFilter('delayed'); setTaskNameFilter('all'); scrollToTable(); }}
              onMouseOver={(e) => e.currentTarget.style.transform = 'scale(1.02)'}
              onMouseOut={(e) => e.currentTarget.style.transform = 'scale(1)'}
            >
              <span className="stat-label flex-center" style={{ justifyContent: 'space-between' }}>
                Delayed
                <AlertCircle size={16} color="var(--color-danger)" />
              </span>
              <span className="stat-value" style={{ color: 'var(--color-danger)' }}>{stats.delayed}</span>
            </div>

            <div
              className="card stat-card area-deadline animate-pulse-red"
              style={{ borderLeft: '4px solid #ef4444', cursor: 'pointer', transition: 'transform 0.2s', position: 'relative' }}
              onClick={() => { setFilter('deadlineNow'); setTaskNameFilter('all'); scrollToTable(); }}
              onMouseOver={(e) => e.currentTarget.style.transform = 'translateY(-2px)'}
              onMouseOut={(e) => e.currentTarget.style.transform = 'translateY(0)'}
            >
              <span className="stat-label flex-center" style={{ justifyContent: 'space-between' }}>
                Deadline Now
                <AlertTriangle size={16} color="#ef4444" />
              </span>
              <span className="stat-value" style={{ color: '#ef4444' }}>{stats.deadlineNow}</span>

              {stats.deadlineNow > 0 && (
                <button
                  onClick={(e) => {
                    e.stopPropagation(); // Prevent card click
                    console.log('Email button clicked');

                    // Filter tasks with Deadline Now
                    const urgentTasks = tasks.filter(t => {
                      const metrics = calculateTimeMetrics(t.status, parseDate(t.startDate), parseDate(t.deadline));
                      return metrics.display === 'Deadline Now';
                    });

                    console.log('Urgent tasks found:', urgentTasks.length);

                    if (urgentTasks.length === 0) {
                      console.log('No urgent tasks to email');
                      return;
                    }

                    const subject = encodeURIComponent(`Urgent: ${urgentTasks.length} Tasks Due Today`);

                    let body = `Hello,\n\nThe following tasks are due today:\n\n`;
                    urgentTasks.forEach(t => {
                      body += `- ${t.name} (Assignee: ${t.assignee || 'Unassigned'})\n`;
                    });
                    body += `\nPlease check them immediately.\n\nBest regards,`;

                    const bodyEncoded = encodeURIComponent(body);
                    const mailtoLink = `mailto:ma.elqorchi@iam.ma?subject=${subject}&body=${bodyEncoded}`;
                    console.log('Opening mailto link:', mailtoLink);
                    window.location.href = mailtoLink;
                  }}
                  style={{
                    marginTop: '0.5rem',
                    padding: '0.25rem 0.5rem',
                    fontSize: '0.75rem',
                    backgroundColor: '#ef4444',
                    color: 'white',
                    border: 'none',
                    borderRadius: '4px',
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '0.25rem',
                    width: 'fit-content'
                  }}
                >
                  <Mail size={12} />
                  Email Report
                </button>
              )}
            </div>

            <div
              className="card stat-card area-completed"
              style={{ borderLeft: '4px solid var(--color-success)', cursor: 'pointer', transition: 'transform 0.2s' }}
              onClick={() => { setFilter('completed'); setTaskNameFilter('all'); scrollToTable(); }}
              onMouseOver={(e) => e.currentTarget.style.transform = 'scale(1.02)'}
              onMouseOut={(e) => e.currentTarget.style.transform = 'scale(1)'}
            >
              <span className="stat-label flex-center" style={{ justifyContent: 'space-between' }}>
                Completed
                <CheckCircle size={16} color="var(--color-success)" />
              </span>
              <span className="stat-value" style={{ color: 'var(--color-success)' }}>{stats.completed}</span>
            </div>

            <div
              className="card stat-card area-inprogress"
              style={{ borderLeft: '4px solid var(--color-warning)', cursor: 'pointer', transition: 'transform 0.2s' }}
              onClick={() => { setFilter('in-progress'); setTaskNameFilter('all'); scrollToTable(); }}
              onMouseOver={(e) => e.currentTarget.style.transform = 'scale(1.02)'}
              onMouseOut={(e) => e.currentTarget.style.transform = 'scale(1)'}
            >
              <span className="stat-label flex-center" style={{ justifyContent: 'space-between' }}>
                In Progress
                <Clock size={16} color="var(--color-warning)" />
              </span>
              <span className="stat-value" style={{ color: 'var(--color-warning)' }}>{stats.inProgress}</span>
            </div>
          </div>

          {/* Task-Specific Dashboards */}
          {Object.keys(taskTypeStats).length > 0 && (
            <>
              <h2 style={{ fontSize: '1.5rem', fontWeight: 700, marginTop: '2rem', marginBottom: '1rem' }}>Task Type Overview</h2>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))', gap: '1.5rem', marginBottom: '2rem' }}>
                {Object.entries(taskTypeStats).map(([taskName, taskStats]) => (
                  <div key={taskName} className="card" style={{ padding: '1.5rem' }}>
                    <div
                      onClick={() => { setTaskNameFilter(taskName); setFilter('all'); scrollToTable(); }}
                      style={{
                        backgroundColor: 'var(--color-bg-secondary)',
                        padding: '0.75rem 1rem',
                        borderRadius: 'var(--radius-md)',
                        marginBottom: '1rem',
                        borderLeft: '4px solid var(--color-primary)',
                        cursor: 'pointer',
                        transition: 'background-color 0.2s'
                      }}
                      onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#e5e7eb'}
                      onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'var(--color-bg-secondary)'}
                    >
                      <h3 style={{
                        fontSize: '1.25rem',
                        fontWeight: 700,
                        color: 'var(--color-text-primary)',
                        margin: 0
                      }}>
                        {taskName}
                      </h3>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', marginBottom: '1rem' }}>
                      {/* Removed redundant total count */}
                      {taskStats.deadlineNow > 0 && (
                        <div className="flex-center animate-pulse-red" style={{
                          color: '#ef4444',
                          fontSize: '0.875rem',
                          fontWeight: '500',
                          padding: '0.25rem 0.5rem',
                          borderRadius: '1rem',
                          backgroundColor: 'rgba(239, 68, 68, 0.1)',
                          gap: '0.25rem',
                          alignSelf: 'flex-start'
                        }}>
                          <AlertTriangle size={14} />
                          {taskStats.deadlineNow} Due Today
                        </div>
                      )}
                    </div>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.75rem' }}>
                      <div
                        onClick={() => { setTaskNameFilter(taskName); setFilter('all'); }}
                        style={{ cursor: 'pointer', transition: 'opacity 0.2s' }}
                        onMouseOver={(e) => e.currentTarget.style.opacity = '0.7'}
                        onMouseOut={(e) => e.currentTarget.style.opacity = '1'}
                      >
                        <div style={{ fontSize: '0.75rem', color: 'var(--color-text-muted)', marginBottom: '0.25rem' }}>Total</div>
                        <div style={{ fontSize: '1.5rem', fontWeight: 700 }}>{taskStats.total}</div>
                      </div>
                      <div
                        onClick={() => { setTaskNameFilter(taskName); setFilter('in-progress'); }}
                        style={{ cursor: 'pointer', transition: 'opacity 0.2s' }}
                        onMouseOver={(e) => e.currentTarget.style.opacity = '0.7'}
                        onMouseOut={(e) => e.currentTarget.style.opacity = '1'}
                      >
                        <div style={{ fontSize: '0.75rem', color: 'var(--color-text-muted)', marginBottom: '0.25rem' }}>In Progress</div>
                        <div style={{ fontSize: '1.5rem', fontWeight: 700, color: 'var(--color-warning)' }}>{taskStats.inProgress}</div>
                      </div>
                      <div
                        onClick={() => { setTaskNameFilter(taskName); setFilter('delayed'); }}
                        style={{ cursor: 'pointer', transition: 'opacity 0.2s' }}
                        onMouseOver={(e) => e.currentTarget.style.opacity = '0.7'}
                        onMouseOut={(e) => e.currentTarget.style.opacity = '1'}
                      >
                        <div style={{ fontSize: '0.75rem', color: 'var(--color-text-muted)', marginBottom: '0.25rem' }}>Delayed</div>
                        <div style={{ fontSize: '1.5rem', fontWeight: 700, color: 'var(--color-danger)' }}>{taskStats.delayed}</div>
                      </div>
                      <div
                        onClick={() => { setTaskNameFilter(taskName); setFilter('completed'); }}
                        style={{ cursor: 'pointer', transition: 'opacity 0.2s' }}
                        onMouseOver={(e) => e.currentTarget.style.opacity = '0.7'}
                        onMouseOut={(e) => e.currentTarget.style.opacity = '1'}
                      >
                        <div style={{ fontSize: '0.75rem', color: 'var(--color-text-muted)', marginBottom: '0.25rem' }}>Completed</div>
                        <div style={{ fontSize: '1.5rem', fontWeight: 700, color: 'var(--color-success)' }}>{taskStats.completed}</div>
                      </div>
                      {/* Only show Reappeared for Interférence Externe */}
                      {(taskName.toLowerCase().includes('interference externe') || taskName.toLowerCase().includes('interférence externe')) && (
                        <div
                          onClick={() => { setTaskNameFilter(taskName); setFilter('reappeared'); }}
                          style={{ cursor: 'pointer', transition: 'opacity 0.2s' }}
                          onMouseOver={(e) => e.currentTarget.style.opacity = '0.7'}
                          onMouseOut={(e) => e.currentTarget.style.opacity = '1'}
                        >
                          <div style={{ fontSize: '0.75rem', color: 'var(--color-text-muted)', marginBottom: '0.25rem' }}>Reappeared</div>
                          <div style={{ fontSize: '1.5rem', fontWeight: 700, color: '#3b82f6' }}>{taskStats.reappeared}</div>
                        </div>
                      )}
                    </div>
                    {/* Progress bar */}
                    <div style={{ marginTop: '1rem', paddingTop: '1rem', borderTop: '1px solid var(--color-border)' }}>
                      <div style={{ fontSize: '0.75rem', color: 'var(--color-text-muted)', marginBottom: '0.5rem' }}>
                        Progress: {taskStats.total > 0 ? Math.round((taskStats.completed / taskStats.total) * 100) : 0}%
                      </div>
                      <div style={{ width: '100%', height: '8px', background: 'var(--color-border)', borderRadius: '4px', overflow: 'hidden' }}>
                        <div style={{
                          width: `${taskStats.total > 0 ? (taskStats.completed / taskStats.total) * 100 : 0}%`,
                          height: '100%',
                          background: 'var(--color-success)',
                          transition: 'width 0.3s ease'
                        }}></div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}

          {/* Controls */}
          <div className="card mb-4" style={{ padding: '1rem', display: 'flex', gap: '1rem', flexWrap: 'wrap', alignItems: 'center' }}>
            <div style={{ position: 'relative', flex: 1, minWidth: '200px' }}>
              <Search size={18} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--color-text-muted)' }} />
              <input
                type="text"
                id="search-tasks-main"
                name="search"
                placeholder="Search tasks or assignees..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                style={{
                  width: '100%',
                  padding: '0.5rem 0.5rem 0.5rem 2.25rem',
                  borderRadius: 'var(--radius-md)',
                  border: '1px solid var(--color-border)',
                  fontSize: '0.875rem'
                }}
              />
            </div>

            {/* Task Name Filter Dropdown */}
            <div style={{ minWidth: '150px' }}>
              <select
                value={taskNameFilter}
                onChange={(e) => setTaskNameFilter(e.target.value)}
                style={{
                  width: '100%',
                  padding: '0.5rem',
                  borderRadius: 'var(--radius-md)',
                  border: '1px solid var(--color-border)',
                  fontSize: '0.875rem',
                  cursor: 'pointer'
                }}
              >
                {uniqueTaskNames.map(name => (
                  <option key={name} value={name}>
                    {name === 'all' ? 'All Tasks' : name}
                  </option>
                ))}
              </select>
            </div>

            <div style={{ display: 'flex', gap: '0.5rem', flexWrap: 'wrap' }}>
              <button
                className={`btn ${filter === 'all' ? 'btn-primary' : 'btn-outline'}`}
                onClick={() => setFilter('all')}
              >
                All Tasks
              </button>
              <button
                className={`btn ${filter === 'deadlineNow' ? 'btn-primary' : 'btn-outline'}`}
                onClick={() => setFilter('deadlineNow')}
                style={filter === 'deadlineNow' ? { backgroundColor: '#ef4444', borderColor: '#ef4444' } : { color: '#ef4444', borderColor: '#ef4444' }}
              >
                <AlertTriangle size={16} />
                Deadline Now
              </button>
              <button
                className={`btn ${filter === 'delayed' ? 'btn-primary' : 'btn-outline'}`}
                onClick={() => setFilter('delayed')}
              >
                Delayed
              </button>
              <button
                className={`btn ${filter === 'completed' ? 'btn-primary' : 'btn-outline'}`}
                onClick={() => setFilter('completed')}
              >
                Completed
              </button>
              <button
                className="btn btn-outline"
                onClick={() => setShowColumnManager(true)}
                title="Manage Columns"
              >
                <Settings size={16} style={{ marginRight: '0.5rem' }} />
                Columns
              </button>
            </div>
          </div>

          {/* Table */}
          <div className="table-container" ref={tableRef}>
            <div className="table-wrapper">
              <table>
                <thead>
                  <tr>
                    {columns.filter(col => !hiddenColumns.includes(col)).map((col, index) => (
                      <th
                        key={index}
                        onClick={() => handleSort(col)}
                        style={{ cursor: 'pointer', userSelect: 'none' }}
                      >
                        <div style={{ display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                          {col}
                          {sortConfig.key === col && (
                            <span>{sortConfig.direction === 'asc' ? '↑' : '↓'}</span>
                          )}
                        </div>
                      </th>
                    ))}
                    <th
                      onClick={() => handleSort('timeMetrics')}
                      style={{ cursor: 'pointer', userSelect: 'none' }}
                    >
                      <div style={{ display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                        Delay / Running
                        {sortConfig.key === 'timeMetrics' && (
                          <span>{sortConfig.direction === 'asc' ? '↑' : '↓'}</span>
                        )}
                      </div>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filteredTasks.map((task) => (
                    <tr
                      key={task.id}
                      onClick={() => handleEditClick(task)}
                      style={{ cursor: 'pointer', transition: 'background-color 0.2s' }}
                      onMouseOver={(e) => e.currentTarget.style.backgroundColor = 'var(--color-bg-secondary)'}
                      onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
                    >
                      {columns.filter(col => !hiddenColumns.includes(col)).map((col, index) => {
                        const val = task[col];
                        const colLower = col.toLowerCase();

                        // Smart rendering based on column name
                        if (colLower.includes('status') || colLower.includes('statut') || colLower.includes('etat')) {
                          // Override status if Deadline is Today
                          if (task.timeMetrics?.display === 'Deadline Now') {
                            return (
                              <td key={index}>
                                <span className="badge badge-red">
                                  Deadline Now
                                </span>
                              </td>
                            );
                          }

                          return (
                            <td key={index}>
                              <span className={`badge ${val === 'Done' || val === 'Completed' || val === 'Terminé' ? 'badge-green' :
                                val === 'In Progress' || val === 'En cours' ? 'badge-yellow' :
                                  val === 'Reappeared' || val === 'Réapparu' ? 'badge-blue' : 'badge-gray'
                                }`}>
                                {val}
                              </span>
                            </td>
                          );
                        }

                        return <td key={index}>{val !== undefined && val !== null ? String(val) : '-'}</td>;
                      })}

                      {/* Calculated Column */}
                      <td>
                        {task.timeMetrics && (
                          <span className={task.timeMetrics.className}>
                            {task.timeMetrics.display}
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                  {filteredTasks.length === 0 && (
                    <tr>
                      <td colSpan={columns.length + 1} className="text-center" style={{ padding: '3rem' }}>
                        No tasks found matching your criteria.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}

      {/* Add Task Modal */}
      {showAddTaskModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0, 0, 0, 0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000
        }} onClick={() => setShowAddTaskModal(false)}>
          <div className="card" style={{
            padding: '2rem',
            maxWidth: '500px',
            width: '90%',
            maxHeight: '90vh',
            overflowY: 'auto'
          }} onClick={(e) => e.stopPropagation()}>
            <h2 style={{ fontSize: '1.5rem', fontWeight: 700, marginBottom: '1.5rem' }}>
              {editingTask ? 'Edit Task' : 'Add New Task'}
            </h2>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
              <div className="form-group">
                <label htmlFor="task-name" style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                  Task Name *
                </label>
                <input
                  type="text"
                  id="task-name"
                  name="name"
                  value={newTask.name}
                  onChange={(e) => setNewTask({ ...newTask, name: e.target.value })}
                  style={{
                    width: '100%',
                    padding: '0.5rem',
                    borderRadius: 'var(--radius-md)',
                    border: '1px solid var(--color-border)',
                    fontSize: '0.875rem'
                  }}
                  placeholder="e.g., ANRT, Interference externe"
                />
              </div>

              <div className="form-group">
                <label htmlFor="task-description" style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                  Description
                </label>
                <input
                  type="text"
                  id="task-description"
                  name="description"
                  value={newTask.description}
                  onChange={(e) => setNewTask({ ...newTask, description: e.target.value })}
                  style={{
                    width: '100%',
                    padding: '0.5rem',
                    borderRadius: 'var(--radius-md)',
                    border: '1px solid var(--color-border)',
                    fontSize: '0.875rem'
                  }}
                  placeholder="Task description"
                />
              </div>

              <div className="form-group">
                <label htmlFor="task-assignee" style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                  Assignee *
                </label>
                <input
                  type="text"
                  id="task-assignee"
                  name="assignee"
                  value={newTask.assignee}
                  onChange={(e) => setNewTask({ ...newTask, assignee: e.target.value })}
                  style={{
                    width: '100%',
                    padding: '0.5rem',
                    borderRadius: 'var(--radius-md)',
                    border: '1px solid var(--color-border)',
                    fontSize: '0.875rem'
                  }}
                  placeholder="Person responsible"
                />
              </div>

              <div className="form-group">
                <label htmlFor="task-status" style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                  Status
                </label>
                <select
                  id="task-status"
                  name="status"
                  value={newTask.status}
                  onChange={(e) => setNewTask({ ...newTask, status: e.target.value })}
                  style={{
                    width: '100%',
                    padding: '0.5rem',
                    borderRadius: 'var(--radius-md)',
                    border: '1px solid var(--color-border)',
                    fontSize: '0.875rem',
                    cursor: 'pointer'
                  }}
                >
                  <option value="In Progress">In Progress</option>
                  <option value="Done">Done</option>
                  <option value="Pending">Pending</option>
                </select>
              </div>

              <div className="form-group">
                <label htmlFor="task-start-date" style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                  Start Date *
                </label>
                <input
                  type="date"
                  id="task-start-date"
                  name="startDate"
                  value={newTask.startDate}
                  onChange={(e) => setNewTask({ ...newTask, startDate: e.target.value })}
                  style={{
                    width: '100%',
                    padding: '0.5rem',
                    borderRadius: 'var(--radius-md)',
                    border: '1px solid var(--color-border)',
                    fontSize: '0.875rem'
                  }}
                />
              </div>

              <div className="form-group">
                <label htmlFor="task-deadline" style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                  Deadline (Optional)
                </label>
                <input
                  type="date"
                  id="task-deadline"
                  name="deadline"
                  value={newTask.deadline}
                  onChange={(e) => setNewTask({ ...newTask, deadline: e.target.value })}
                  style={{
                    width: '100%',
                    padding: '0.5rem',
                    borderRadius: 'var(--radius-md)',
                    border: '1px solid var(--color-border)',
                    fontSize: '0.875rem'
                  }}
                />
              </div>

              {/* Custom Fields */}
              {customFields.map((field, index) => (
                <div key={index}>
                  <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                    {field.name || 'Custom Field'}
                  </label>
                  <input
                    type={field.type === 'date' ? 'date' : field.type === 'number' ? 'number' : 'text'}
                    value={newTask[`custom_${index}`] || ''}
                    onChange={(e) => setNewTask({ ...newTask, [`custom_${index}`]: e.target.value })}
                    style={{
                      width: '100%',
                      padding: '0.5rem',
                      borderRadius: 'var(--radius-md)',
                      border: '1px solid var(--color-border)',
                      fontSize: '0.875rem'
                    }}
                    placeholder={field.name}
                  />
                </div>
              ))}

              <div style={{ display: 'flex', gap: '0.5rem', marginTop: '1rem' }}>
                {editingTask ? (
                  <>
                    <button className="btn btn-primary" onClick={handleUpdateTask} style={{ flex: 1 }}>
                      Update Task
                    </button>
                    <button className="btn btn-danger" onClick={handleDeleteTask} style={{ flex: 1, background: 'var(--color-danger)', color: 'white', border: 'none' }}>
                      Delete
                    </button>
                  </>
                ) : (
                  <button className="btn btn-primary" onClick={handleAddTask} style={{ flex: 1 }}>
                    Add Task
                  </button>
                )}
                <button className="btn btn-outline" onClick={() => {
                  setShowAddTaskModal(false);
                  setEditingTask(null);
                  setNewTask({ name: '', description: '', assignee: '', status: 'In Progress', startDate: '', deadline: '' });
                }} style={{ flex: 1 }}>
                  Cancel
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Export Modal */}
      {showExportModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0, 0, 0, 0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000
        }} onClick={() => setShowExportModal(false)}>
          <div className="card" style={{
            padding: '2rem',
            maxWidth: '400px',
            width: '90%'
          }} onClick={(e) => e.stopPropagation()}>
            <h2 style={{ fontSize: '1.5rem', fontWeight: 700, marginBottom: '1.5rem' }}>Export Data</h2>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>

              {/* Scope Selection */}
              <div>
                <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                  Data to Export
                </label>
                <div style={{ display: 'flex', gap: '1rem' }}>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', cursor: 'pointer' }}>
                    <input
                      type="radio"
                      name="scope"
                      checked={exportFilters.scope === 'all'}
                      onChange={() => setExportFilters({ ...exportFilters, scope: 'all' })}
                    />
                    All Data
                  </label>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', cursor: 'pointer' }}>
                    <input
                      type="radio"
                      name="scope"
                      checked={exportFilters.scope === 'filtered'}
                      onChange={() => setExportFilters({ ...exportFilters, scope: 'filtered' })}
                    />
                    Filtered Data Only
                  </label>
                </div>
              </div>

              {/* Format Selection */}
              <div>
                <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                  Format
                </label>
                <div style={{ display: 'flex', gap: '1rem' }}>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', cursor: 'pointer' }}>
                    <input
                      type="radio"
                      name="format"
                      checked={exportFilters.format === 'excel'}
                      onChange={() => setExportFilters({ ...exportFilters, format: 'excel' })}
                    />
                    Excel (.xlsx)
                  </label>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', cursor: 'pointer' }}>
                    <input
                      type="radio"
                      name="format"
                      checked={exportFilters.format === 'csv'}
                      onChange={() => setExportFilters({ ...exportFilters, format: 'csv' })}
                    />
                    CSV (.csv)
                  </label>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', cursor: 'pointer' }}>
                    <input
                      type="radio"
                      name="format"
                      checked={exportFilters.format === 'numbers'}
                      onChange={() => setExportFilters({ ...exportFilters, format: 'numbers' })}
                    />
                    Numbers (Mac)
                  </label>
                </div>
              </div>

              {/* Legacy Filters (Only show if Scope is All) */}
              {exportFilters.scope === 'all' && (
                <div style={{ padding: '1rem', background: 'var(--color-bg-secondary)', borderRadius: 'var(--radius-md)' }}>
                  <div style={{ marginBottom: '1rem' }}>
                    <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                      Filter by Task Name (Optional)
                    </label>
                    <select
                      value={exportFilters.taskName}
                      onChange={(e) => setExportFilters({ ...exportFilters, taskName: e.target.value })}
                      style={{
                        width: '100%',
                        padding: '0.5rem',
                        borderRadius: 'var(--radius-md)',
                        border: '1px solid var(--color-border)',
                        fontSize: '0.875rem'
                      }}
                    >
                      {uniqueTaskNames.map(name => (
                        <option key={name} value={name}>
                          {name === 'all' ? 'All Tasks' : name}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                    <input
                      type="checkbox"
                      id="delayedOnly"
                      checked={exportFilters.delayedOnly}
                      onChange={(e) => setExportFilters({ ...exportFilters, delayedOnly: e.target.checked })}
                      style={{ cursor: 'pointer' }}
                    />
                    <label htmlFor="delayedOnly" style={{ fontSize: '0.875rem', cursor: 'pointer' }}>
                      Only export delayed tasks
                    </label>
                  </div>
                </div>
              )}

              <div style={{ display: 'flex', gap: '0.5rem', marginTop: '0.5rem' }}>
                <button
                  className="btn btn-primary"
                  onClick={() => {
                    handleExportToExcel();
                    setShowExportModal(false);
                  }}
                  style={{ flex: 1 }}
                >
                  Export
                </button>
                <button className="btn btn-outline" onClick={() => setShowExportModal(false)} style={{ flex: 1 }}>
                  Cancel
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
      {/* Settings Modal */}
      {showSettingsModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0, 0, 0, 0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000
        }} onClick={() => setShowSettingsModal(false)}>
          <div className="card" style={{
            padding: '2rem',
            maxWidth: '500px',
            width: '90%'
          }} onClick={(e) => e.stopPropagation()}>
            <h2 style={{ fontSize: '1.5rem', fontWeight: 700, marginBottom: '1.5rem' }}>Airtable Settings</h2>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>


              <div>
                <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                  Personal Access Token (API Key)
                </label>
                <input
                  type="password"
                  value={airtableConfig.apiKey}
                  onChange={(e) => {
                    const newVal = e.target.value;
                    setAirtableConfig({ ...airtableConfig, apiKey: newVal });
                    localStorage.setItem('at_apiKey', newVal);
                  }}
                  style={{
                    width: '100%',
                    padding: '0.5rem',
                    borderRadius: 'var(--radius-md)',
                    border: '1px solid var(--color-border)',
                    fontSize: '0.875rem'
                  }}
                  placeholder="pat..."
                />
              </div>

              <div>
                <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                  Base ID
                </label>
                <input
                  type="text"
                  value={airtableConfig.baseId}
                  onChange={(e) => {
                    const newVal = e.target.value;
                    setAirtableConfig({ ...airtableConfig, baseId: newVal });
                    localStorage.setItem('at_baseId', newVal);
                  }}
                  style={{
                    width: '100%',
                    padding: '0.5rem',
                    borderRadius: 'var(--radius-md)',
                    border: '1px solid var(--color-border)',
                    fontSize: '0.875rem'
                  }}
                  placeholder="app..."
                />
              </div>

              <div>
                <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.875rem', fontWeight: 500 }}>
                  Table Name
                </label>
                <input
                  type="text"
                  value={airtableConfig.tableName}
                  onChange={(e) => {
                    const newVal = e.target.value;
                    setAirtableConfig({ ...airtableConfig, tableName: newVal });
                    localStorage.setItem('at_tableName', newVal);
                  }}
                  style={{
                    width: '100%',
                    padding: '0.5rem',
                    borderRadius: 'var(--radius-md)',
                    border: '1px solid var(--color-border)',
                    fontSize: '0.875rem'
                  }}
                  placeholder="Tasks"
                />
              </div>



              <div style={{ display: 'flex', gap: '0.5rem', marginTop: '1.5rem' }}>
                <button className="btn btn-primary" onClick={() => setShowSettingsModal(false)} style={{ flex: 1 }}>
                  Save & Close
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
      {/* Column Manager Modal */}
      {showColumnManager && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0, 0, 0, 0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000
        }} onClick={() => setShowColumnManager(false)}>
          <div className="card" style={{
            padding: '2rem',
            maxWidth: '400px',
            width: '90%',
            maxHeight: '80vh',
            overflowY: 'auto'
          }} onClick={(e) => e.stopPropagation()}>
            <h2 style={{ fontSize: '1.5rem', fontWeight: 700, marginBottom: '1.5rem' }}>Manage Columns</h2>
            <p style={{ fontSize: '0.875rem', color: 'var(--color-text-muted)', marginBottom: '1rem' }}>
              Check to show, uncheck to hide. Use arrows to reorder.
            </p>

            {hiddenColumns.length > 0 && (
              <button
                onClick={() => {
                  if (window.confirm(`Are you sure you want to permanently delete ${hiddenColumns.length} hidden columns?`)) {
                    const newCols = columns.filter(col => !hiddenColumns.includes(col));
                    setColumns(newCols);
                    setHiddenColumns([]);
                    localStorage.setItem('tableColumns', JSON.stringify(newCols));
                    localStorage.setItem('hiddenColumns', JSON.stringify([]));
                  }
                }}
                className="btn"
                style={{
                  width: '100%',
                  marginBottom: '1rem',
                  backgroundColor: '#fee2e2',
                  color: '#ef4444',
                  border: '1px solid #fecaca',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '0.5rem'
                }}
              >
                <Trash2 size={16} />
                Delete {hiddenColumns.length} Hidden Columns
              </button>
            )}

            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
              {columns.map((col, index) => (
                <div key={col} style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  padding: '0.5rem',
                  border: '1px solid var(--color-border)',
                  borderRadius: 'var(--radius-md)',
                  background: 'var(--color-bg-secondary)'
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                    <input
                      type="checkbox"
                      checked={!hiddenColumns.includes(col)}
                      onChange={(e) => {
                        const newHidden = e.target.checked
                          ? hiddenColumns.filter(c => c !== col)
                          : [...hiddenColumns, col];
                        setHiddenColumns(newHidden);
                        localStorage.setItem('hiddenColumns', JSON.stringify(newHidden));
                      }}
                      style={{ cursor: 'pointer' }}
                    />
                    <span style={{ fontSize: '0.875rem', fontWeight: 500 }}>{col}</span>
                  </div>
                  <div style={{ display: 'flex', gap: '0.25rem' }}>
                    <button
                      disabled={index === 0}
                      onClick={() => {
                        const newCols = [...columns];
                        [newCols[index - 1], newCols[index]] = [newCols[index], newCols[index - 1]];
                        setColumns(newCols);
                        localStorage.setItem('tableColumns', JSON.stringify(newCols));
                      }}
                      style={{
                        background: 'none', border: 'none', cursor: index === 0 ? 'default' : 'pointer',
                        opacity: index === 0 ? 0.3 : 1, padding: '0.25rem'
                      }}
                    >
                      ⬆️
                    </button>
                    <button
                      disabled={index === columns.length - 1}
                      onClick={() => {
                        const newCols = [...columns];
                        [newCols[index + 1], newCols[index]] = [newCols[index], newCols[index + 1]];
                        setColumns(newCols);
                        localStorage.setItem('tableColumns', JSON.stringify(newCols));
                      }}
                      style={{
                        background: 'none', border: 'none', cursor: index === columns.length - 1 ? 'default' : 'pointer',
                        opacity: index === columns.length - 1 ? 0.3 : 1, padding: '0.25rem'
                      }}
                    >
                      ⬇️
                    </button>
                  </div>
                </div>
              ))}
            </div>

            <div style={{ marginTop: '1.5rem', paddingTop: '1rem', borderTop: '1px solid var(--color-border)' }}>
              <h3 style={{ fontSize: '1rem', fontWeight: 600, marginBottom: '0.75rem' }}>Add New Column</h3>
              <div style={{ display: 'flex', gap: '0.5rem', marginBottom: '0.5rem' }}>
                <input
                  type="text"
                  placeholder="Column Name"
                  value={newColumnName}
                  onChange={(e) => setNewColumnName(e.target.value)}
                  style={{
                    flex: 1,
                    padding: '0.5rem',
                    borderRadius: 'var(--radius-md)',
                    border: '1px solid var(--color-border)',
                    fontSize: '0.875rem'
                  }}
                />
                <select
                  value={newColumnType}
                  onChange={(e) => setNewColumnType(e.target.value)}
                  style={{
                    padding: '0.5rem',
                    borderRadius: 'var(--radius-md)',
                    border: '1px solid var(--color-border)',
                    fontSize: '0.875rem'
                  }}
                >
                  <option value="text">Text</option>
                  <option value="number">Number</option>
                  <option value="date">Date</option>
                </select>
              </div>
              <button
                className="btn btn-outline"
                style={{ width: '100%' }}
                onClick={() => {
                  if (!newColumnName.trim()) {
                    alert('Please enter a column name');
                    return;
                  }
                  if (columns.includes(newColumnName)) {
                    alert('Column already exists');
                    return;
                  }

                  // Add to custom fields
                  const updatedCustomFields = [...customFields, { name: newColumnName, type: newColumnType }];
                  setCustomFields(updatedCustomFields);
                  localStorage.setItem('customFields', JSON.stringify(updatedCustomFields));

                  // Add to columns list
                  const updatedColumns = [...columns, newColumnName];
                  setColumns(updatedColumns);
                  localStorage.setItem('tableColumns', JSON.stringify(updatedColumns));

                  setNewColumnName('');
                  setNewColumnType('text');
                }}
              >
                + Add Column
              </button>
            </div>

            <div style={{ marginTop: '1.5rem' }}>
              <button className="btn btn-primary" onClick={() => setShowColumnManager(false)} style={{ width: '100%' }}>
                Done
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
