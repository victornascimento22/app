import React, { useCallback } from 'react';
import { format, isValid, parse } from 'date-fns';
import { ptBR } from 'date-fns/locale';
import { Calendar as CalendarIcon } from 'lucide-react';
import { Button } from "@/app/components/ui/button";
import { Calendar } from "@/app/components/ui/calendar";
import { Label } from "@/app/components/ui/label";
import { Input } from "@/app/components/ui/input";

// Componente DateSelector que recebe várias props para manipulação de data
export function DateSelector({
  label,
  date,
  setDate,
  dateInput,
  setDateInput,
  isOpen,
  setIsOpen
}: {
  label: string;
  date: Date | null;
  setDate: (date: Date | null) => void;
  dateInput: string;
  setDateInput: (input: string) => void;
  isOpen: boolean;
  setIsOpen: (isOpen: boolean) => void;
}) {
  // Função para lidar com a mudança de data no input
  const handleDateChange = useCallback((
    value: string
  ) => {
    const formattedValue = formatDateInput(value);
    setDateInput(formattedValue);
   
    // Verifica se o valor formatado tem 10 caracteres (formato completo de data)
    if (formattedValue.length === 10) {
      const parsedDate = parse(formattedValue, 'dd/MM/yyyy', new Date());
      // Se a data for válida, atualiza a data; caso contrário, define como null
      if (isValid(parsedDate)) {
        setDate(parsedDate);
      } else {
        setDate(null);
      }
    } else {
      setDate(null);
    }
  }, [setDate, setDateInput]);

  // Função para lidar com a seleção de data no calendário
  const handleCalendarSelect = useCallback((
    selectedDate: Date | undefined
  ) => {
    if (selectedDate) {
      setDate(selectedDate);
      setDateInput(format(selectedDate, 'dd/MM/yyyy'));
      setIsOpen(false); // Fecha o calendário após a seleção
    }
  }, [setDate, setDateInput, setIsOpen]);

  // Função para formatar a entrada de data
  const formatDateInput = useCallback((input: string): string => {
    const cleaned = input.replace(/\D/g, ''); // Remove caracteres não numéricos
    const match = cleaned.match(/^(\d{0,2})(\d{0,2})(\d{0,4})$/);
   
    if (!match) return '';
   
    const [, day, month, year] = match;
   
    let formatted = '';
    if (day) formatted += day;
    if (month) formatted += `/${month}`;
    if (year) formatted += `/${year}`;
   
    return formatted;
  }, []);

  return (
    <div>
      <Label htmlFor={label} className="text-white mb-2 block">{label}</Label>
      <div className="relative">
        <Input
          type="text"
          id={label}
          name={label}
          value={dateInput}
          onChange={(e) => handleDateChange(e.target.value)}
          placeholder="DD/MM/YYYY"
          className="w-full pr-10 bg-white border-gray-60 text-black"
          maxLength={10}
        />
        <Button
          type="button"
          variant="ghost"
          size="icon"
          className="absolute right-0 top-0 h-full"
          onClick={() => setIsOpen(!isOpen)} // Alterna a visibilidade do calendário
        >
          <CalendarIcon className="h-4 w-4 text-gray-400" />
        </Button>
        {isOpen && (
          <div className="absolute top-full left-0 w-full mt-1 z-10">
            <Calendar
              mode="single"
              selected={date || undefined}
              onSelect={handleCalendarSelect}
              locale={ptBR}
              defaultMonth={date || undefined}
              className="border border-gray-600 bg-white text-BLACK rounded-md shadow-lg"
              classNames={{
                day_selected: "border border-gray-600 bg-gray-700 rounded-md shadow-lg",
                day_today: "bg-gray-600 text-white",
              }}
            />
          </div>
        )}
      </div>
    </div>
  );
}
