import React, { createContext, useContext, useReducer, useEffect, ReactNode } from 'react';
import { AppView, OfficeHostType } from '../types';
import { detectOfficeHost } from '../services/office/OfficeContext';

export interface AppState {
  view: AppView;
  host: OfficeHostType;
  error: string | null;
  notification: { type: 'success' | 'error' | 'warning' | 'info'; message: string } | null;
  isLoading: boolean;
  pendingPrompt: string | null;
}

type Action =
  | { type: 'SET_VIEW'; payload: AppView }
  | { type: 'SET_OFFICE_CONTEXT'; payload: OfficeHostType }
  | { type: 'SET_ERROR'; payload: string | null }
  | { type: 'SET_NOTIFICATION'; payload: { type: 'success' | 'error' | 'warning' | 'info'; message: string } | null }
  | { type: 'CLEAR_NOTIFICATION' }
  | { type: 'SET_LOADING'; payload: boolean }
  | { type: 'SET_PENDING_PROMPT'; payload: string | null };

const initialState: AppState = {
  view: 'chat',
  host: 'Unknown',
  error: null,
  notification: null,
  isLoading: false,
  pendingPrompt: null,
};

const AppContext = createContext<{ state: AppState; dispatch: React.Dispatch<Action> } | undefined>(undefined);

function appReducer(state: AppState, action: Action): AppState {
  switch (action.type) {
    case 'SET_VIEW': return { ...state, view: action.payload };
    case 'SET_OFFICE_CONTEXT': return { ...state, host: action.payload };
    case 'SET_ERROR': return { ...state, error: action.payload };
    case 'SET_NOTIFICATION': return { ...state, notification: action.payload };
    case 'CLEAR_NOTIFICATION': return { ...state, notification: null };
    case 'SET_LOADING': return { ...state, isLoading: action.payload };
    case 'SET_PENDING_PROMPT': return { ...state, pendingPrompt: action.payload };
    default: return state;
  }
}

export const AppProvider = ({ children }: { children: ReactNode }) => {
  const [state, dispatch] = useReducer(appReducer, initialState);

  // Detect Office host on mount
  useEffect(() => {
    const host = detectOfficeHost();
    dispatch({ type: 'SET_OFFICE_CONTEXT', payload: host });
  }, []);

  // Listen for localStorage quota errors dispatched by storage.ts
  // so the user gets a visible warning instead of silently losing data.
  useEffect(() => {
    const onStorageError = (e: Event) => {
      const detail = (e as CustomEvent).detail;
      dispatch({
        type: 'SET_NOTIFICATION',
        payload: {
          type: 'warning',
          message: 'Storage full — older conversations may not be saved. Try clearing old chats.',
        },
      });
      console.warn('[Storage Error]', detail);
    };
    window.addEventListener('storage-error', onStorageError as EventListener);
    return () => window.removeEventListener('storage-error', onStorageError as EventListener);
  }, []);

  // Auto-clear notifications after 4 seconds
  useEffect(() => {
    if (state.notification) {
      const timer = setTimeout(() => dispatch({ type: 'CLEAR_NOTIFICATION' }), 4000);
      return () => clearTimeout(timer);
    }
  }, [state.notification]);

  return (
    <AppContext.Provider value={{ state, dispatch }}>
      {children}
    </AppContext.Provider>
  );
};

export const useAppState = () => {
  const context = useContext(AppContext);
  if (!context) throw new Error('useAppState must be used within AppProvider');
  return context.state;
};

export const useAppDispatch = () => {
  const context = useContext(AppContext);
  if (!context) throw new Error('useAppDispatch must be used within AppProvider');
  return context.dispatch;
};
