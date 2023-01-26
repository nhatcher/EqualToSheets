import { useEffect, useState, useRef } from 'react';
import { getModule, CalcModule } from './modelLoader';

export type { CalcModule } from './modelLoader';

export const useCalcModule = () => {
  const modulePromise = useRef<Promise<CalcModule> | null>(null);
  const [calcModule, setCalcModule] = useState<CalcModule | null>(null);
  useEffect(() => {
    if (modulePromise.current === null) {
      modulePromise.current = getModule();
      modulePromise.current.then((module) => {
        setCalcModule(module);
      });
    }
  }, []);
  return { calcModule };
};
