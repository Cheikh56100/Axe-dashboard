import React from 'react';

interface LandingPageProps {
  onGetStarted: () => void;
  onLogin: () => void;
}

export function LandingPage({ onGetStarted, onLogin }: LandingPageProps) {
  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 font-sans selection:bg-amber-500 selection:text-slate-950">
      {/* Navigation */}
      <nav className="flex items-center justify-between px-8 py-6 max-w-7xl mx-auto border-b border-slate-800/80">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-xl bg-amber-500 flex items-center justify-center font-bold text-slate-950 shadow-lg shadow-amber-500/20">
            OFX
          </div>
          <span className="font-bold text-xl tracking-tight text-white">OFX Bridge</span>
        </div>
        <div className="flex items-center gap-4">
          <button 
            onClick={onLogin}
            className="text-sm font-medium text-slate-300 hover:text-white transition-colors px-4 py-2"
          >
            Se connecter
          </button>
          <button 
            onClick={onGetStarted}
            className="text-sm font-semibold bg-amber-500 hover:bg-amber-400 text-slate-950 px-5 py-2.5 rounded-lg transition-all shadow-md shadow-amber-500/10 active:scale-95"
          >
            Créer un compte
          </button>
        </div>
      </nav>

      {/* Hero Section */}
      <section className="max-w-5xl mx-auto text-center px-6 pt-20 pb-16">
        <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full border border-amber-500/30 bg-amber-500/10 text-amber-400 text-xs font-medium mb-8">
          <span className="w-2 h-2 rounded-full bg-amber-400 animate-pulse" />
          Conversion instantanée & sécurisée
        </div>
        <h1 className="text-4xl sm:text-6xl font-extrabold text-white tracking-tight mb-6 leading-tight">
          Vos relevés PDF, transformés en <span className="text-amber-400">données financières</span>
        </h1>
        <p className="text-lg text-slate-400 max-w-2xl mx-auto mb-10 leading-relaxed">
          Importez vos fichiers bancaires PDF et vérifiez vos transactions en quelques secondes. Obtenez un export OFX ou Excel propre et prêt à l'emploi.
        </p>
        <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
          <button 
            onClick={onGetStarted}
            className="w-full sm:w-auto px-8 py-4 bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold rounded-xl shadow-lg shadow-amber-500/20 transition-all text-base active:scale-95"
          >
            Commencer gratuitement
          </button>
        </div>
      </section>

      {/* Features - Bento Grid */}
      <section className="max-w-6xl mx-auto px-6 py-12">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="bg-slate-900/60 border border-slate-800 p-8 rounded-2xl md:col-span-2 hover:border-slate-700 transition-colors">
            <div className="w-10 h-10 rounded-lg bg-amber-500/10 border border-amber-500/20 text-amber-400 flex items-center justify-center font-bold mb-4">
              ⚡
            </div>
            <h3 className="text-xl font-bold text-white mb-2">Traitement ultra-rapide</h3>
            <p className="text-slate-400 text-sm leading-relaxed">
              Extraction instantanée des relevés PDF sans altérer la précision des montants et des dates.
            </p>
          </div>

          <div className="bg-slate-900/60 border border-slate-800 p-8 rounded-2xl hover:border-slate-700 transition-colors">
            <div className="w-10 h-10 rounded-lg bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 flex items-center justify-center font-bold mb-4">
              🔒
            </div>
            <h3 className="text-xl font-bold text-white mb-2">Données sécurisées</h3>
            <p className="text-slate-400 text-sm leading-relaxed">
              Isolement complet de vos données bancaires avec chiffrement et protection RLS.
            </p>
          </div>

          <div className="bg-slate-900/60 border border-slate-800 p-8 rounded-2xl hover:border-slate-700 transition-colors">
            <div className="w-10 h-10 rounded-lg bg-blue-500/10 border border-blue-500/20 text-blue-400 flex items-center justify-center font-bold mb-4">
              📊
            </div>
            <h3 className="text-xl font-bold text-white mb-2">Exports multiples</h3>
            <p className="text-slate-400 text-sm leading-relaxed">
              Exportation instantanée au format OFX pour vos logiciels comptables ou Excel.
            </p>
          </div>

          <div className="bg-slate-900/60 border border-slate-800 p-8 rounded-2xl md:col-span-2 hover:border-slate-700 transition-colors">
            <div className="w-10 h-10 rounded-lg bg-purple-500/10 border border-purple-500/20 text-purple-400 flex items-center justify-center font-bold mb-4">
              🎯
            </div>
            <h3 className="text-xl font-bold text-white mb-2">Vérification intégrée</h3>
            <p className="text-slate-400 text-sm leading-relaxed">
              Contrôlez et modifiez facilement les transactions extraites avant la conversion finale.
            </p>
          </div>
        </div>
      </section>
    </div>
  );
}
