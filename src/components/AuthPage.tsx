import React, { useState } from 'react';
import { supabase } from '../lib/supabase';
import { FileSpreadsheet, Mail, Lock, UserRound, AlertCircle, Loader2 } from 'lucide-react';

export default function AuthPage() {
  const [mode, setMode] = useState<'login' | 'signup'>('login');
  const [name, setName] = useState('');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setSuccessMsg(null);
    setLoading(true);

    try {
      if (mode === 'login') {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
      } else {
        const { error } = await supabase.auth.signUp({
          email,
          password,
          options: { data: { name: name.trim() } },
        });
        if (error) throw error;
        setSuccessMsg('Cadastro realizado! Verifique seu e-mail para confirmar a conta.');
      }
    } catch (err: any) {
      const msg = err.message || 'Erro desconhecido';
      if (msg.includes('Invalid login credentials')) setError('E-mail ou senha incorretos.');
      else if (msg.includes('User already registered')) setError('Este e-mail já está cadastrado.');
      else if (msg.includes('Password should be at least')) setError('Senha deve ter pelo menos 6 caracteres.');
      else setError(msg);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#0F0E1A] text-[#F4F5FC] flex items-center justify-center p-4">
      <div className="absolute top-0 left-0 right-0 h-96 bg-gradient-to-b from-[#252344]/15 to-transparent pointer-events-none" />
      <div className="absolute top-1/4 right-0 w-96 h-96 bg-[#7C3AED]/5 rounded-full filter blur-[120px] pointer-events-none" />
      <div className="absolute bottom-1/4 left-0 w-96 h-96 bg-[#6366F1]/5 rounded-full filter blur-[120px] pointer-events-none" />

      <div className="relative w-full max-w-sm">
        <div className="flex flex-col items-center mb-8">
          <div className="h-14 w-14 rounded-2xl bg-gradient-to-tr from-[#7C3AED] to-[#6366F1] flex items-center justify-center shadow-lg shadow-[#7C3AED]/20 mb-4">
            <FileSpreadsheet className="h-7 w-7 text-white" />
          </div>
          <h1 className="text-xl font-bold tracking-tight text-[#F4F5FC]">Conversor de Planilhas Fiscal</h1>
          <span className="text-[11px] bg-[#1A1830] text-[#8B80FF] px-2 py-0.5 rounded-full border border-[#2E2C4A] mt-1">Uso Interno</span>
        </div>

        <div className="bg-[#141229] border border-[#2E2C4A] rounded-3xl p-7 shadow-2xl">
          <div className="flex bg-[#0F0E1A] p-1 rounded-xl border border-[#2E2C4A] mb-6">
            <button
              onClick={() => { setMode('login'); setError(null); setSuccessMsg(null); }}
              className={`flex-1 py-1.5 text-xs font-bold rounded-lg transition-all ${mode === 'login' ? 'bg-[#6366F1] text-white shadow' : 'text-[#A1A6C4] hover:text-white'}`}
            >
              Entrar
            </button>
            <button
              onClick={() => { setMode('signup'); setError(null); setSuccessMsg(null); }}
              className={`flex-1 py-1.5 text-xs font-bold rounded-lg transition-all ${mode === 'signup' ? 'bg-[#6366F1] text-white shadow' : 'text-[#A1A6C4] hover:text-white'}`}
            >
              Criar Conta
            </button>
          </div>

          <form onSubmit={handleSubmit} className="space-y-4">
            {mode === 'signup' && (
              <div>
                <label className="block text-[11px] font-semibold text-[#A1A6C4] mb-1.5">Nome</label>
                <div className="relative">
                  <UserRound className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-[#A1A6C4]" />
                  <input
                    type="text"
                    value={name}
                    onChange={(e) => setName(e.target.value)}
                    required
                    placeholder="Seu nome completo"
                    className="w-full bg-[#0F0E1A] border border-[#2E2C4A] focus:border-[#6366F1] rounded-xl pl-9 pr-3 py-2.5 text-sm outline-none text-[#F4F5FC] placeholder-[#A1A6C4]/30 transition-colors"
                  />
                </div>
              </div>
            )}

            <div>
              <label className="block text-[11px] font-semibold text-[#A1A6C4] mb-1.5">E-mail</label>
              <div className="relative">
                <Mail className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-[#A1A6C4]" />
                <input
                  type="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  required
                  placeholder="seu@email.com"
                  className="w-full bg-[#0F0E1A] border border-[#2E2C4A] focus:border-[#6366F1] rounded-xl pl-9 pr-3 py-2.5 text-sm outline-none text-[#F4F5FC] placeholder-[#A1A6C4]/30 transition-colors"
                />
              </div>
            </div>

            <div>
              <label className="block text-[11px] font-semibold text-[#A1A6C4] mb-1.5">Senha</label>
              <div className="relative">
                <Lock className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-[#A1A6C4]" />
                <input
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  required
                  placeholder="••••••••"
                  minLength={6}
                  className="w-full bg-[#0F0E1A] border border-[#2E2C4A] focus:border-[#6366F1] rounded-xl pl-9 pr-3 py-2.5 text-sm outline-none text-[#F4F5FC] placeholder-[#A1A6C4]/30 transition-colors"
                />
              </div>
            </div>

            {error && (
              <div className="flex items-start gap-2 bg-[#FF3D00]/10 border border-[#FF3D00]/30 text-[#FF3D00] rounded-xl p-3 text-xs">
                <AlertCircle className="h-4 w-4 shrink-0 mt-0.5" />
                {error}
              </div>
            )}

            {successMsg && (
              <div className="flex items-start gap-2 bg-[#22C55E]/10 border border-[#22C55E]/30 text-[#22C55E] rounded-xl p-3 text-xs">
                <AlertCircle className="h-4 w-4 shrink-0 mt-0.5" />
                {successMsg}
              </div>
            )}

            <button
              type="submit"
              disabled={loading}
              className="w-full py-2.5 bg-gradient-to-r from-[#6366F1] to-[#7C3AED] text-white text-sm font-bold rounded-xl shadow shadow-indigo-500/20 hover:shadow-indigo-500/40 transition-all disabled:opacity-60 disabled:cursor-not-allowed flex items-center justify-center gap-2"
            >
              {loading && <Loader2 className="h-4 w-4 animate-spin" />}
              {mode === 'login' ? 'Entrar' : 'Criar Conta'}
            </button>
          </form>
        </div>
      </div>
    </div>
  );
}
