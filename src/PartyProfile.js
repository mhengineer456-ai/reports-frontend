import React from 'react';

function PartyProfile({ parties, partyData, setPartyData, onSubmit, onBack }) {
  return (
    <div className="component-container">
      <h2>Party Profile Management</h2>
      <div className="form-card">
        <form onSubmit={onSubmit}>
          <div className="form-group">
            <label>Party/Business Name *</label>
            <input 
              type="text" 
              placeholder="Enter party or business name"
              value={partyData.name}
              onChange={(e) => setPartyData({...partyData, name: e.target.value})}
              required
            />
          </div>
          <div className="form-group">
            <label>Contact Person Name</label>
            <input 
              type="text" 
              placeholder="Enter contact person name"
              value={partyData.contactPerson}
              onChange={(e) => setPartyData({...partyData, contactPerson: e.target.value})}
            />
          </div>
          <div className="form-group">
            <label>Phone Number *</label>
            <input 
              type="tel" 
              placeholder="Enter contact number"
              value={partyData.contact}
              onChange={(e) => setPartyData({...partyData, contact: e.target.value})}
              required
            />
          </div>
          <div className="form-group">
            <label>Email Address</label>
            <input 
              type="email" 
              placeholder="Enter email address"
              value={partyData.email}
              onChange={(e) => setPartyData({...partyData, email: e.target.value})}
            />
          </div>
          <div className="form-group">
            <label>Complete Address</label>
            <textarea 
              placeholder="Enter complete address with PIN code"
              value={partyData.address}
              onChange={(e) => setPartyData({...partyData, address: e.target.value})}
              rows="3"
            />
          </div>
          <div className="form-group">
            <label>GST Number (if applicable)</label>
            <input 
              type="text" 
              placeholder="Enter GST number"
              value={partyData.gst}
              onChange={(e) => setPartyData({...partyData, gst: e.target.value})}
            />
          </div>
          <div className="form-actions">
            <button type="submit" className="save-btn">Save Party Profile</button>
          </div>
        </form>
      </div>
      
      {parties.length > 0 && (
        <div className="party-list-full">
          <h3>Registered Parties ({parties.length})</h3>
          <div className="party-grid">
            {parties.map(party => (
              <div key={party.id} className="party-card">
                <h4>{party.name}</h4>
                <p><strong>Contact:</strong> {party.contact}</p>
                <p><strong>Email:</strong> {party.email || 'N/A'}</p>
                <p><strong>GST:</strong> {party.gst || 'N/A'}</p>
                <p><strong>Address:</strong> {party.address?.substring(0, 50)}...</p>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

export default PartyProfile;