import { Alert, AlertDescription, AlertTitle } from "@/app/components/ui/alert"
import { ExclamationTriangleIcon } from "@radix-ui/react-icons"

interface WarnProps {
  title: string;
  messages: string[];
}

export function Warn({ title, messages }: WarnProps) {
  return (
    <Alert variant="destructive">
      <ExclamationTriangleIcon className="h-4 w-4" />
      <AlertTitle>{title}</AlertTitle>
      <AlertDescription>
        <ul className="list-disc pl-5">
          {messages.map((message, index) => (
            <li key={index}>{message}</li>
          ))}
        </ul>
      </AlertDescription>
    </Alert>
  )
}
