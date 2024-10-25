import { format } from 'date-fns';

interface ReportParams {
  refExt: string;
  selectedTypes: number[];
  startDate: Date | null;
  endDate: Date | null;
}

export async function generateHiltiReport({
  refExt,
  selectedTypes,
  startDate,
  endDate,
}: ReportParams): Promise<Blob> {
  const baseUrl = 'http://172.16.14.94:5000/';
  const params = new URLSearchParams();

  params.append('ref_ext', refExt.trim());
  params.append('tipo_nota', selectedTypes.join(','));
  if (startDate) params.append('data_de', format(startDate, 'dd/MM/yyyy'));
  if (endDate) params.append('data_ate', format(endDate, 'dd/MM/yyyy'));

  const url = `${baseUrl}/reportHilti?${params.toString()}`;

  console.log('Attempting to fetch from URL:', url);
  console.log('Request Parameters:', Object.fromEntries(params));

  const response = await fetch(url, {
    method: 'GET',
    mode: 'cors',
    headers: {
      'Accept': 'application/json, application/xlsx',
      'Content-Type': 'application/json',
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('Error response:', errorText);
    throw new Error(`Erro do servidor: ${response.status} - ${errorText || response.statusText}`);
  }

  return await response.blob();
}
