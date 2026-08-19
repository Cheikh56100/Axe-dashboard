import React, { useState } from 'react';
import { supabase } from '../lib/supabase';

interface AuthProps {
  initialMode?: 'login' | 'signup';
  onSuccess: () => void;
  onCancel: () => void;
}

export function AuthModal({ initialMode = 'login', onSuccess, onCancel }: AuthProps) {
  const [isLogin, setIsLogin] = useState(initialMode === 'login');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setLoading(true);

    try {
      if (isLogin) {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
      } else {
        const { error } = await supabase.auth.signUp({ email, password });
        if (error) throw error;
      }
      onSuccess();
    } catch (err: any) {
      setError(err.message || 'Une erreur est survenue');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4 z-50">
      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-8 max-w-md w-full shadow-2xl">
        <h2 className="text-2xl font-bold text-white mb-2">
          {isLogin ? 'Connexion' : 'Créer un compte'}
        </h2>
        <p className="text-slate-400 text-sm mb-6">
          {isLogin ? 'Accédez à votre espace sécurisé' : 'Rejoignez-nous pour convertir vos relevés'}
        </p>

        {error && (
          <div className="bg-red-500/10 border border-red-500/20 text-red-400 p-3 rounded-lg text-sm mb-4">
            {error}
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-xs font-semibold text-slate-300 uppercase mb-2">Email</label>
            <input
              type="email"
              required
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-lg px-4 py-3 text-white focus:outline-none focus:border-amber-500 text-sm"
              placeholder="nom@exemple.com"
            />
          </div>
          <div>
            <label className="block text-xs font-semibold text-slate-300 uppercase mb-2">Mot de passe</label>
            <input
              type="password"
              required
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-lg px-4 py-3 text-white focus:outline-none focus:border-amber-500 text-sm"
              placeholder="••••••••"
            />
          </div>

          <button
            type="submit"
            disabled={loading}
            className="w-full bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold py-3 rounded-lg transition-colors text-sm disabled:opacity-50"
          >
            {loading ? 'Chargement...' : isLogin ? 'Se connecter' : "S'inscrire"}
          </button>
        </form>

        <div className="mt-6 text-center text-sm text-slate-400">
          {isLogin ? (
            <p>
              Pas encore de compte ?{' '}
              <button onClick={() => setIsLogin(false)} className="text-amber-400 font-semibold hover:underline">
                S'inscrire
              </button>
            </p>
          ) : (
            <p>
              Déjà un compte ?{' '}
              <button onClick={() => setIsLogin(true)} className="text-amber-400 font-semibold hover:underline">
                Se connecter
              </button>
            </p>
          )}
        </div>

        <button onClick={onCancel} className="mt-4 text-xs text-slate-500 w-full hover:text-slate-300">
          Fermer
        </button>
      </div>
    </div>
  );
}
