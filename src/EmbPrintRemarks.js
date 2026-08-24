import React, { useState } from 'react';
import { useHistory } from 'react-router-dom';
import EmbroideryChallan from './EmbroideryChallan';
import PrintingChallan from './PrintingChallan';

const EmbPrintRemarks = () => {
  const history = useHistory();
  const [activeTab, setActiveTab] = useState('embroidery'); // 'embroidery' | 'printing'

  const styles = {
    container: {
      minHeight: '100vh',
      backgroundColor: '#f8fafc',
      fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
      paddingBottom: '40px'
    },
    header: {
      background: 'linear-gradient(135deg, #1e1b4b 0%, #312e81 60%, #4338ca 100%)',
      color: '#ffffff',
      padding: '24px 32px',
      boxShadow: '0 10px 25px -5px rgba(49, 46, 129, 0.3)',
      marginBottom: '16px'
    },
    headerTop: {
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'space-between',
      marginBottom: '16px'
    },
    backButton: {
      display: 'inline-flex',
      alignItems: 'center',
      gap: '8px',
      background: 'rgba(255, 255, 255, 0.12)',
      color: '#ffffff',
      border: '1px solid rgba(255, 255, 255, 0.2)',
      padding: '8px 18px',
      borderRadius: '10px',
      fontSize: '0.88rem',
      fontWeight: '600',
      cursor: 'pointer',
      transition: 'all 0.2s ease'
    },
    titleSection: {
      display: 'flex',
      flexDirection: 'column',
      gap: '4px'
    },
    mainTitle: {
      fontSize: '1.75rem',
      fontWeight: '800',
      letterSpacing: '-0.02em',
      margin: 0
    },
    subTitle: {
      fontSize: '0.92rem',
      color: '#c7d2fe',
      margin: 0
    },
    tabsWrapper: {
      display: 'flex',
      gap: '12px',
      marginTop: '20px'
    },
    tabButton: (isActive) => ({
      display: 'inline-flex',
      alignItems: 'center',
      gap: '10px',
      padding: '12px 28px',
      borderRadius: '12px',
      fontSize: '0.95rem',
      fontWeight: '700',
      cursor: 'pointer',
      border: 'none',
      transition: 'all 0.25s cubic-bezier(0.4, 0, 0.2, 1)',
      background: isActive ? '#ffffff' : 'rgba(255, 255, 255, 0.15)',
      color: isActive ? '#1e1b4b' : '#ffffff',
      boxShadow: isActive ? '0 8px 20px rgba(0, 0, 0, 0.15)' : 'none'
    })
  };

  return (
    <div style={styles.container}>
      {/* Top Navigation Banner */}
      <header style={styles.header}>
        <div style={styles.headerTop}>
          <button style={styles.backButton} onClick={() => history.push('/dashboard')}>
            ← Back to Dashboard
          </button>
        </div>

        <div style={styles.titleSection}>
          <h1 style={styles.mainTitle}>EMB / PRINT REMARKS TRACKER</h1>
          <p style={styles.subTitle}>
            Real-time pending lots & remarks updates for Embroidery and Printing production
          </p>
        </div>

        {/* Tab Switchers */}
        <div style={styles.tabsWrapper}>
          <button
            style={styles.tabButton(activeTab === 'embroidery')}
            onClick={() => setActiveTab('embroidery')}
          >
            <span>🧵</span> EMBROIDERY REMARKS (Pending Lots)
          </button>

          <button
            style={styles.tabButton(activeTab === 'printing')}
            onClick={() => setActiveTab('printing')}
          >
            <span>🖨️</span> PRINTING REMARKS (Pending Lots)
          </button>
        </div>
      </header>

      {/* Render Component Replica depending on active tab */}
      <div>
        {activeTab === 'embroidery' ? (
          <EmbroideryChallan initialEmbFilter="pending" isEmbedded={true} />
        ) : (
          <PrintingChallan initialPrintingStatusFilter="pending" isEmbedded={true} />
        )}
      </div>
    </div>
  );
};

export default EmbPrintRemarks;
