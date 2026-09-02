import test from 'node:test';
import assert from 'node:assert/strict';

import {
  extractDropboxReferencePlan,
  selectDropboxTaskFiles,
} from '../task-dropbox-routing.js';

const failedTaskInstruction = `Retrieve the tenant information from one of the PDF files in the same Dropbox folder and insert it into the report.

File: Tenant_Apartment_Condition_Report.docx
Path: /LONDON - ACCESS/MINA CAPITAL/4_5165_DES SOURCES/Tenant_Apartment_Condition_Report.docx`;

test('derives the containing Dropbox folder from the exact failed task instruction', () => {
  assert.deepEqual(extractDropboxReferencePlan(failedTaskInstruction), {
    referenced: true,
    explicitFiles: [
      '/LONDON - ACCESS/MINA CAPITAL/4_5165_DES SOURCES/Tenant_Apartment_Condition_Report.docx',
    ],
    folders: ['/LONDON - ACCESS/MINA CAPITAL/4_5165_DES SOURCES'],
  });
});

test('selects the referenced DOCX and every PDF requested from the same folder', () => {
  const entries = [
    {
      type: 'file',
      name: 'Tenant_Apartment_Condition_Report.docx',
      path: '/LONDON - ACCESS/MINA CAPITAL/4_5165_DES SOURCES/Tenant_Apartment_Condition_Report.docx',
      size: 1000,
    },
    {
      type: 'file',
      name: 'Noureddine 5165#4.pdf',
      path: '/LONDON - ACCESS/MINA CAPITAL/4_5165_DES SOURCES/Noureddine 5165#4.pdf',
      size: 2000,
    },
    {
      type: 'file',
      name: '4-5165 Rules.pdf',
      path: '/LONDON - ACCESS/MINA CAPITAL/4_5165_DES SOURCES/4-5165 Rules.pdf',
      size: 3000,
    },
    {
      type: 'file',
      name: 'building-photo.jpg',
      path: '/LONDON - ACCESS/MINA CAPITAL/4_5165_DES SOURCES/building-photo.jpg',
      size: 4000,
    },
  ];

  const selected = selectDropboxTaskFiles({
    entries,
    explicitFiles: extractDropboxReferencePlan(failedTaskInstruction).explicitFiles,
    instruction: failedTaskInstruction,
    maxFiles: 10,
  });

  assert.deepEqual(
    selected.map((item) => item.name),
    [
      'Tenant_Apartment_Condition_Report.docx',
      'Noureddine 5165#4.pdf',
      '4-5165 Rules.pdf',
    ]
  );
});

test('recognizes escaped underscores and HTML spaces from copied task email text', () => {
  const copied =
    'Path: /LONDON - ACCESS/MINA CAPITAL/4\\_5165\\_DES&#x20; SOURCES/Tenant_Apartment_Condition_Report.docx';

  assert.deepEqual(extractDropboxReferencePlan(copied).folders, [
    '/LONDON - ACCESS/MINA CAPITAL/4_5165_DES SOURCES',
  ]);
});

test('joins a Dropbox path wrapped across email lines', () => {
  const wrapped = `Path: /LONDON - ACCESS/MINA CAPITAL/4_5165_DES
SOURCES/Tenant_Apartment_Condition_Report.docx`;

  assert.deepEqual(extractDropboxReferencePlan(wrapped).explicitFiles, [
    '/LONDON - ACCESS/MINA CAPITAL/4_5165_DES SOURCES/Tenant_Apartment_Condition_Report.docx',
  ]);
});
