const fs = require('fs');
const path = require('path');

const logPath = '/Users/ant/Library/Logs/Claude/mcp-server-excel.log';
const content = fs.readFileSync(logPath, 'utf8');
const lines = content.split('\n');
const lastLines = lines.slice(-100).join('\n');
console.log(lastLines);
