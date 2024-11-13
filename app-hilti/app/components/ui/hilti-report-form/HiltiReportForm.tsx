'use client'

import React, { useState, useCallback, useMemo } from 'react'
import { format } from 'date-fns'
import { ptBR } from 'date-fns/locale'
import { RefExtInput } from './RefExtInput'
import { NoteTypeSelector } from './NoteTypeSelector'
import { DateSelector } from './DateSelector'
import { GenerateReportButton } from './GenerateReportButton'
import { generateHiltiReport } from '@/app/api/api'
import { Warn } from './Warn'

export default function HiltiReportForm() {
  const [refExt, setRefExt] = useState<string>('')
  const [selectedTypes, setSelectedTypes] = useState<number[]>([])
  const [startDate, setStartDate] = useState<Date | null>(null)
  const [endDate, setEndDate] = useState<Date | null>(null)
  const [startDateInput, setStartDateInput] = useState('')
  const [endDateInput, setEndDateInput] = useState('')
  const [isStartDateOpen, setIsStartDateOpen] = useState(false)
  const [isEndDateOpen, setIsEndDateOpen] = useState(false)
  const [isLoading, setIsLoading] = useState<boolean>(false)
  const [warnings, setWarnings] = useState<string[]>([])

  const noteTypes = useMemo(() => [
    { id: 0, value: 'comum', label: 'Comum' },
    { id: 3, value: 'complementar', label: 'Complementar' },
    { id: 5, value: 'corretiva', label: 'Corretiva' },
    { id: 1, value: 'retorno', label: 'Retorno' },
    { id: 9, value: 'outros_doc', label: 'Outros Documentos' },
  ], [])

  //lógica para gerar o relatório
  const handleGenerateReport = async () => {
    setIsLoading(true)
    try {
      const reportData = await generateHiltiReport({
        refExt,
        selectedTypes: selectedTypes.map(Number),
        startDate,
        endDate
      })
      if (selectedTypes.every(item => typeof item == 'number' && !isNaN(item))) {
        console.log('Tipos de nota enviados:', selectedTypes.map(Number));
      }else{
        console.log('Tipos de nota enviados:', selectedTypes);
      }
      
      const url = window.URL.createObjectURL(reportData)
      const a = document.createElement('a')
      a.href = url
      a.download = 'relatorio.xlsx'
      a.click()
      a.remove()
      window.URL.revokeObjectURL(url)
      console.log('Relatório gerado com sucesso:', reportData)
    } catch (error) {
      console.error('Erro ao gerar relatório:', error)
    } finally {
      setIsLoading(false)
    }
  }

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setWarnings([]);

    // Validações
    const newWarnings = [];
    if (!refExt.trim()) newWarnings.push("O campo Ref. Ext. é obrigatório.");
    if (selectedTypes.length === 0) newWarnings.push("Selecione pelo menos um tipo de nota.");
    if (!startDate || !endDate) newWarnings.push("Selecione um intervalo de datas válido.");

    if (newWarnings.length > 0) {
      setWarnings(newWarnings);
      return;
    }

    try {
      const blob = await generateHiltiReport({
        refExt,
        selectedTypes,
        startDate: startDate || null,
        endDate: endDate || null,
      });
      // Lógica para lidar com o blob (por exemplo, fazer o download)
    } catch (error) {
      setWarnings(["Ocorreu um erro ao gerar o relatório. Por favor, tente novamente."]);
    }
  };

  return (
    <div className="w-screen min-h-screen overflow-auto bg-fixed bg-cover bg-no-repeat flex justify-center items-center py-12" style={{backgroundImage: "url('https://www.capitaltrade.srv.br/wp-content/uploads/2024/10/background_CT_2024-1.png')"}}>
      <form onSubmit={handleSubmit} className="w-full max-w-2xl bg-gray-800 bg-opacity-50 p-8 rounded-lg shadow-lg font-panton">
        <h1 className="text-center text-3xl font-panton text-white mb-8">RELATÓRIO HILTI</h1>
       
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
          <RefExtInput value={refExt} onChange={(e) => setRefExt(e.target.value)} />
          <NoteTypeSelector 
            selectedTypes={selectedTypes} 
            setSelectedTypes={(types) => setSelectedTypes(types)} 
            noteTypes={noteTypes} 
          />
          <DateSelector
            label="Data Inicial"
            date={startDate}
            setDate={setStartDate}
            dateInput={startDateInput}
            setDateInput={setStartDateInput}
            isOpen={isStartDateOpen}
            setIsOpen={setIsStartDateOpen}
          />
          <DateSelector
            label="Data Final"
            date={endDate}
            setDate={setEndDate}
            dateInput={endDateInput}
            setDateInput={setEndDateInput}
            isOpen={isEndDateOpen}
            setIsOpen={setIsEndDateOpen}
          />
        </div>
 
        <div className="flex justify-center w-full">
          <GenerateReportButton onClick={handleGenerateReport} isLoading={isLoading} />
        </div>
      </form>
    </div>
  )
}
