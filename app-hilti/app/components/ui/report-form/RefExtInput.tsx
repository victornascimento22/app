import React from 'react';
import { Label } from "@/app/components/ui/label";
import { Input } from "@/app/components/ui/input";

export function RefExtInput({ value, onChange }: { value: string; onChange: (e: React.ChangeEvent<HTMLInputElement>) => void }) {
  return (
    <div className="col-span-1 md:col-span-2">
      <Label htmlFor="ref_ext" className="text-white mb-2 block ">
        Ref. Externa
      </Label>
      <Input
        type="text"
        id="ref_ext"
        value={value}
        onChange={onChange}
        className="w-full bg-white border-white text-black"
        placeholder="Referência Externa"
      />
    </div>
  );
}
