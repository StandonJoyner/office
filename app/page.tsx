'use client';

import { useMemo } from 'react';
import { WordEditor } from '@/ui/word';
import { TraceabilityPanel } from '@/ui/common';
import { DataUpdateNotification } from '@/ui/components/DataUpdateNotification';
import { createTraceabilityService } from '@/core/trace';

export default function Home() {
  const traceabilityService = useMemo(
    () => createTraceabilityService(
      () => ({ type: 'doc', content: [] }),
      () => []
    ),
    []
  );
  return (
    <main className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 p-8">
      {/* Data Update Notification */}
      <DataUpdateNotification />

      <div className="mx-auto max-w-4xl">
        <h1 className="text-4xl font-bold text-slate-900">
          Office Suite
        </h1>
        <p className="mt-4 text-lg text-slate-600">
          Word & Excel integrated office suite with data traceability
        </p>

        <div className="mt-8 grid gap-6 md:grid-cols-2">
          <div className="col-span-1">
            <h2 className="text-2xl font-semibold text-slate-900 mb-4">Word Editor</h2>
            <WordEditor
              content={{
                type: 'doc',
                content: [
                  {
                    type: 'paragraph',
                    content: [
                      { type: 'text', text: 'Welcome to the Office Suite!' },
                      { type: 'text', text: 'This document demonstrates data references.' },
                    ],
                  },
                ],
              }}
              onChange={(content) => console.log('Content changed:', content)}
            />
          </div>

          <div className="col-span-1">
            <h2 className="text-2xl font-semibold text-slate-900 mb-4">Data Traceability</h2>
            <TraceabilityPanel documentId="demo-doc-1" traceabilityService={traceabilityService} />
          </div>
        </div>

        <div className="mt-8 grid gap-6 md:grid-cols-2">
          <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
            <h2 className="text-xl font-semibold text-slate-900">Features</h2>
            <ul className="mt-2 list-inside list-disc text-slate-600">
              <li>Rich text editing</li>
              <li>Data references to Excel cells</li>
              <li>Reference state indicators</li>
              <li>Click to navigate to source</li>
              <li>Manual refresh of references</li>
            </ul>
          </div>
          <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
            <h2 className="text-xl font-semibold text-slate-900">Traceability</h2>
            <ul className="mt-2 list-inside list-disc text-slate-600">
              <li>Version snapshots</li>
              <li>Change history tracking</li>
              <li>Dependency analysis</li>
              <li>Audit logging</li>
            </ul>
          </div>
        </div>
      </div>
    </main>
  )
}
