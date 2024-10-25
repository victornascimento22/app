import React, { useState, useCallback } from 'react'
import { ChevronDown } from 'lucide-react'
import { Button } from "@/app/components/ui/button"
import { Label } from "@/app/components/ui/label"
import { Checkbox } from "@/app/components/ui/checkbox"

export function NoteTypeSelector({ selectedTypes, setSelectedTypes, noteTypes }: { selectedTypes: number[]; setSelectedTypes: (types: number[]) => void; noteTypes: { id: number; value: string; label: string }[] }) {
  const [isDropdownOpen, setIsDropdownOpen] = useState(false);

  const toggleDropdown = useCallback(() => setIsDropdownOpen(!isDropdownOpen), [isDropdownOpen]);

  const handleTypeChange = useCallback((checked: boolean, value: number) => {
    setSelectedTypes(checked
      ? [...selectedTypes, value]
      : selectedTypes.filter((type: number) => type !== value)
    );
  }, [selectedTypes, setSelectedTypes]);

  const handleSelectAll = useCallback((checked: boolean) => {
    setSelectedTypes(checked ? noteTypes.map(type => type.id) : []);
  }, [noteTypes, setSelectedTypes]);

  return (
    <div className="col-span-1 md:col-span-2">
      <Label className="text-white mb-2 block">
        Tipo de Nota
      </Label>
      <Button
        type="button"
        onClick={toggleDropdown}
        className="w-full justify-between bg-white text-black hover:bg-orange-500 0 mb-2"
      >
        <span>{selectedTypes.length ? `${selectedTypes.length} selecionado(s)` : 'Tipo de Nota'}</span>
        <ChevronDown className="h-4 w-4 " />
      </Button>
      {isDropdownOpen && (
        <div className="bg-white rounded-lg shadow p-3 space-y-2 mt-2">
          <div className="flex items-center">
            <Checkbox
              id="select-all"
              checked={selectedTypes.length === noteTypes.length}
              onCheckedChange={handleSelectAll}
              className='border-black text-white hover:bg-blue-800'
            />
            <Label htmlFor="select-all" className="ml-2 text-sm font-medium text-black">
              Selecionar Todos
            </Label>
          </div>
          {noteTypes.map((type) => (
            <div key={type.value} className="flex items-center">
              <Checkbox
                id={type.value}
                checked={selectedTypes.includes(type.id)}
                onCheckedChange={(checked: boolean) => handleTypeChange(checked, type.id)}
                className="border-black text-blue-500 focus:ring-blue-500"
              />
              <Label htmlFor={type.value} className="ml-2 text-sm font-medium text-black">
                {type.label}
              </Label>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
