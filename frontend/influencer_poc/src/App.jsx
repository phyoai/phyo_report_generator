import { useState, useEffect } from 'react';
import './App.css';

function App() {
  const apiUrl = import.meta.env.VITE_API_URL || 'http://localhost:5000';
  const proxiedMediaUrl = (url) => url ? `${apiUrl}/api/proxy-media?url=${encodeURIComponent(url)}` : '';
  const formatBudget = (financial) => {
    const amount = financial?.totalBudget;
    if (!amount && amount !== 0) return '';
    const currency = financial?.budgetCurrency || '';
    return currency ? `${amount} ${currency}` : `${amount}`;
  };
  // Set body background to prevent white flash
  useEffect(() => {
    document.body.style.backgroundColor = 'var(--bg-canvas)';
    document.documentElement.style.backgroundColor = 'var(--bg-canvas)';
    return () => {
      document.body.style.backgroundColor = '';
      document.documentElement.style.backgroundColor = '';
    };
  }, []);
  const [prompt, setPrompt] = useState('');
  const [instagramPostUrl, setInstagramPostUrl] = useState('');
  const [images, setImages] = useState([]);
  const [previews, setPreviews] = useState([]);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);

  useEffect(() => {
    return () => {
      previews.forEach((url) => URL.revokeObjectURL(url));
    };
  }, [previews]);

  const handleImageUpload = (e) => {
    const files = Array.from(e.target.files);
    
    // Add new files to existing images
    const newImages = [...images, ...files];
    setImages(newImages);
    
    // Create previews for all images
    const newPreviews = newImages.map(f => URL.createObjectURL(f));
    setPreviews(newPreviews);
    
    // Clear the input so same file can be re-uploaded
    e.target.value = '';
  };

  const removeImage = (index) => {
    const newImages = images.filter((_, i) => i !== index);
    const newPreviews = previews.filter((_, i) => i !== index);
    setImages(newImages);
    setPreviews(newPreviews);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    
    if (!prompt.trim() && images.length === 0 && !instagramPostUrl.trim()) {
      setError('Please provide a prompt, images, or a post URL');
      return;
    }

    setLoading(true);
    setError(null);
    setResult(null);

    try {
      const formData = new FormData();
      formData.append('prompt', prompt);
      if (instagramPostUrl.trim()) {
        const url = instagramPostUrl.trim();
        if (url.includes('youtube.com') || url.includes('youtu.be')) {
          formData.append('youtube_post_url', url);
        } else {
          formData.append('instagram_post_url', url);
        }
      }

      // Append each image individually
      images.forEach((img, index) => {
        formData.append('images', img);
        console.log(`Adding image ${index + 1}:`, img.name, img.type, img.size);
      });

      console.log('Sending request with', images.length, 'images');

      const res = await fetch(`${apiUrl}/api/generate-report`, {
        method: 'POST',
        body: formData,
      });

      const data = await res.json();

      if (data.success) {
        setResult(data);
      } else {
        setError(data.error || 'Failed to generate report');
      }
    } catch (err) {
      setError('Error: ' + err.message);
    } finally {
      setLoading(false);
    }
  };

  // const downloadReport = () => {
  //   if (result?.filename) {
  //     window.open(`http://localhost:5000/api/download/${result.filename}`, '_blank');
  //   }
  // };

  const resetForm = () => {
    setPrompt('');
    setInstagramPostUrl('');
    setImages([]);
    setPreviews([]);
    setResult(null);
    setError(null);
  };

  return (
    <div className="report-app min-h-screen w-full relative overflow-x-hidden">
      {/* Background pattern */}
      <div className="report-pattern fixed inset-0 w-full h-full opacity-50 pointer-events-none">
        <div className="absolute inset-0 w-full h-full transform rotate-45"></div>
        <div className="absolute inset-0 w-full h-full"></div>
      </div>
      
      <div className="w-full min-h-screen flex flex-col relative z-10">
        <div className="container mx-auto px-3 sm:px-4 py-4 sm:py-6 md:py-8 flex-1">
          <div className="max-w-5xl mx-auto">
            {/* Glassmorphism container */}
            <div className="report-shell backdrop-blur-xl rounded-2xl sm:rounded-3xl p-4 sm:p-6 md:p-8 lg:p-10">
            
            {/* Header */}
            <div className="report-hero mb-6 sm:mb-8 md:mb-10">
              <h1 className="report-title report-heading">
                Campaign Report
                <span className="report-heading-accent">Generator</span>
              </h1>
              <p className="report-subtitle report-lead text-sm sm:text-base md:text-lg leading-relaxed px-2">
                Turn campaign screenshots, reels, and dashboards into polished presentation-ready reports in minutes.
              </p>
            </div>

            {/* Form */}
            <form onSubmit={handleSubmit} className="space-y-4 sm:space-y-6 md:space-y-8">
              
              {/* Prompt Input */}
              <div className="report-section space-y-2 sm:space-y-3">
                <div className="report-section-title">Campaign Brief</div>
                <label className="report-label block text-sm sm:text-base font-semibold mb-2">
                  Campaign Description
                </label>
                <textarea
                  value={prompt}
                  onChange={(e) => setPrompt(e.target.value)}
                  placeholder="Example: Create report for summer fashion campaign with Influencer X, budget $5000, ran June 1-30. Include Instagram and YouTube metrics."
                  rows="3"
                  disabled={loading}
                  className="report-input w-full px-3 py-2 sm:px-4 sm:py-3 rounded-lg sm:rounded-xl text-sm sm:text-base resize-none disabled:opacity-50 disabled:cursor-not-allowed transition-all duration-200"
                />
              </div>

              {/* Instagram / YouTube Post URL */}
              <div className="report-section space-y-2">
                <div className="report-section-title">Content Source</div>
                <label className="report-label block text-sm sm:text-base font-semibold mb-1">
                  Post / Reel / YouTube Video URL
                </label>
                <div className="report-dropzone flex items-center rounded-lg sm:rounded-xl overflow-hidden transition-all duration-200">
                  <span className="pl-3 report-muted select-none">
                    <svg className="w-5 h-5" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1" />
                    </svg>
                  </span>
                  <input
                    type="url"
                    value={instagramPostUrl}
                    onChange={(e) => setInstagramPostUrl(e.target.value)}
                    placeholder="https://www.instagram.com/reel/ABC123/"
                    disabled={loading}
                    className="flex-1 px-3 py-2 sm:py-3 bg-transparent report-title text-sm sm:text-base focus:outline-none disabled:opacity-50 disabled:cursor-not-allowed"
                  />
                </div>
                <p className="report-faint text-xs">Paste the link to the reel, post, or YouTube video to auto-fetch exact metrics (likes, views, comments)</p>
              </div>

              {/* Image Upload */}
              <div className="report-section space-y-2 sm:space-y-3">
                <div className="report-section-title">Visual Inputs</div>
                <label className="report-label block text-sm sm:text-base font-semibold mb-2">
                  Upload Screenshots ({images.length} selected)
                </label>
                {/* What to upload guide */}
                {/* <div className="report-alert-info backdrop-blur-md rounded-lg p-3 mb-2">
                  <p className="token-info text-xs font-semibold mb-2">What to upload for exact metrics:</p>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 text-xs report-subtitle">
                    <div className="flex gap-2 items-start">
                      <span className="token-primary text-base">1.</span>
                      <div>
                        <p className="font-semibold report-title">Instagram Insights</p>
                        <p>Open Reel -> tap <span className="token-primary-soft font-mono">...</span> -> <span className="token-primary-soft">View Insights</span> -> screenshot</p>
                        <p className="report-faint mt-0.5">Shows: Plays, Reach, Likes, Saves, Shares</p>
                      </div>
                    </div>
                    <div className="flex gap-2 items-start">
                      <span className="token-error text-base">2.</span>
                      <div>
                        <p className="font-semibold report-title">YouTube Analytics</p>
                        <p>YouTube Studio -> Video -> <span className="token-error">Analytics</span> -> screenshot</p>
                        <p className="report-faint mt-0.5">Shows: Views, Watch time, CTR</p>
                      </div>
                    </div>
                    <div className="flex gap-2 items-start">
                      <span className="token-accent text-base">3.</span>
                      <div>
                        <p className="font-semibold report-title">Campaign dashboard</p>
                        <p>Any tool showing campaign numbers (Cloutflow, Phyllo, etc.)</p>
                        <p className="report-faint mt-0.5">All visible numbers extracted automatically</p>
                      </div>
                    </div>
                  </div>
                </div> */}
                <div className="relative">
                  <input
                    type="file"
                    multiple
                    accept="image/*"
                    onChange={handleImageUpload}
                    disabled={loading}
                    className="report-dropzone w-full px-3 py-4 sm:px-4 sm:py-6 rounded-lg sm:rounded-xl cursor-pointer disabled:cursor-not-allowed disabled:opacity-50 transition-all duration-200 report-title text-sm sm:text-base file:mr-2 sm:file:mr-4 file:py-1 sm:file:py-2 file:px-2 sm:file:px-4 file:rounded file:border-0 file:text-xs sm:file:text-sm file:font-semibold file:bg-[var(--brand-primary)] file:text-[var(--brand-ink)] hover:file:bg-[var(--brand-secondary)]"
                  />
                  <div className="report-upload-copy">
                    <div className="report-upload-icon">
                      <svg className="w-5 h-5" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 0115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                      </svg>
                    </div>
                    <div className="report-upload-title">Drop files here or browse</div>
                    <div className="report-upload-subtitle">Insights screenshots, creator content, logo, and dashboards</div>
                  </div>
                </div>
                <p className="report-subtitle text-xs sm:text-sm mt-1 sm:mt-2 leading-relaxed">
                  Upload campaign photos, metrics screenshots, brand logo, and creator content. You can select multiple files at once or add more later.
                </p>
              </div>

              {/* Image Previews */}
              {previews.length > 0 && (
                <div className="report-panel backdrop-blur-md p-3 sm:p-4 md:p-6 rounded-lg sm:rounded-xl">
                  <div className="flex flex-col sm:flex-row sm:justify-between sm:items-center mb-3 sm:mb-4 gap-2">
                    <h3 className="font-semibold report-title text-sm sm:text-base md:text-lg">
                      Uploaded Images ({previews.length})
                    </h3>
                    <button
                      type="button"
                      onClick={() => {
                        setImages([]);
                        setPreviews([]);
                      }}
                      className="text-xs sm:text-sm font-medium transition-colors duration-200 self-start sm:self-auto text-[var(--brand-secondary)] hover:text-[var(--text-primary)]"
                    >
                      Clear All
                    </button>
                  </div>
                  <div className="report-preview-grid">
                    {previews.map((url, idx) => (
                      <div key={idx} className="report-preview-card relative group">
                        <img
                          src={url}
                          alt={`Preview ${idx + 1}`}
                          className="w-full h-16 sm:h-20 md:h-24 object-cover rounded border border-[var(--border-subtle)] group-hover:border-[var(--brand-secondary)] transition-all duration-200"
                        />
                        <span className="absolute bottom-0.5 left-0.5 sm:bottom-1 sm:left-1 report-chip backdrop-blur-sm text-[var(--text-primary)] text-xs px-1.5 py-0.5 rounded text-center">
                          {idx + 1}
                        </span>
                        <button
                          type="button"
                          onClick={() => removeImage(idx)}
                          className="absolute top-0.5 right-0.5 sm:top-1 sm:right-1 bg-[var(--brand-primary)] backdrop-blur-sm text-[var(--brand-ink)] rounded-full w-5 h-5 sm:w-6 sm:h-6 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-all duration-200 text-xs"
                          title="Remove image"
                        >
                          ×
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Submit Button */}
              <button
                type="submit"
                disabled={loading || (!prompt.trim() && images.length === 0 && !instagramPostUrl.trim())}
                className="report-button-primary w-full font-semibold py-3 sm:py-4 px-4 sm:px-6 rounded-lg sm:rounded-xl disabled:opacity-50 disabled:cursor-not-allowed transition-all duration-300 flex items-center justify-center gap-2 sm:gap-3 shadow-lg hover:shadow-xl backdrop-blur-sm transform hover:scale-[1.01] sm:hover:scale-[1.02] active:scale-[0.99] sm:active:scale-[0.98]"
              >
                {loading ? (
                  <>
                    <svg className="animate-spin h-4 w-4 sm:h-5 sm:w-5 text-[var(--brand-ink)]" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                    </svg>
                    <span className="text-sm sm:text-base md:text-lg">Generating Report...</span>
                  </>
                ) : (
                  <span className="text-sm sm:text-base md:text-lg">Generate Report</span>
                )}
              </button>
            </form>

            {/* Error Alert */}
            {error && (
              <div className="report-alert-danger mt-4 sm:mt-6 backdrop-blur-md p-3 sm:p-4 rounded-lg sm:rounded-xl">
                <div className="flex flex-col sm:flex-row sm:items-center gap-1 sm:gap-0">
                  <span className="font-semibold text-sm sm:text-base">Error: </span>
                  <span className="sm:ml-2 text-sm sm:text-base break-words">{error}</span>
                </div>
              </div>
            )}

            {/* Success Result */}
            {result && (
              <div className="mt-4 sm:mt-6 space-y-3 sm:space-y-4 md:space-y-6">
                <div className="report-alert-success backdrop-blur-md p-3 sm:p-4 rounded-lg sm:rounded-xl">
                  <p className="font-semibold text-sm sm:text-base md:text-lg">
                    Report generated successfully!
                  </p>
                </div>

                {/* Warnings — e.g. Instagram URL blocked */}
                {result.warnings && result.warnings.length > 0 && (
                  <div className="space-y-2">
                    {result.warnings.map((w, i) => (
                      <div key={i} className="report-alert-warning backdrop-blur-md p-3 sm:p-4 rounded-lg sm:rounded-xl flex gap-3">
                        <span className="token-warning text-lg shrink-0">!</span>
                        <div>
                          <p className="text-sm sm:text-base">{w}</p>
                          <p className="text-xs mt-1 opacity-80">
                            <strong>How to fix:</strong> Go to Instagram → open the Reel → tap "View Insights" → screenshot that screen → upload it here.
                          </p>
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                {/* Extracted Data */}
                <div className="report-panel backdrop-blur-md p-3 sm:p-4 md:p-6 rounded-lg sm:rounded-xl">
                  <h2 className="report-title text-lg sm:text-xl md:text-2xl font-bold mb-3 sm:mb-4 md:mb-6">
                    Extracted Data
                  </h2>

                  {/* helper: show value only if non-empty and non-zero */}
                  {(() => {
                    const d = result.extracted_data;
                    const val = (v) => (!v || v === '0' || v === 0) ? '' : v;

                    return (
                  <div className="space-y-3 sm:space-y-4 md:space-y-6">
                    {/* Campaign Info */}
                    <div className="report-panel-strong backdrop-blur-sm p-3 sm:p-4 rounded-lg sm:rounded-xl border-l-4 border-l-[var(--brand-primary)]">
                      <h3 className="report-title font-bold text-base sm:text-lg md:text-xl mb-1 sm:mb-2">
                        {val(d.campaignName) || 'Campaign Report'}
                      </h3>
                      <div className="flex flex-wrap gap-x-4 gap-y-1 text-sm report-subtitle">
                        {val(d.brand)    && <span>Brand: <strong className="report-title">{d.brand}</strong></span>}
                        {val(d.creator)  && <span>Creator: <strong className="report-title">{d.creator}</strong></span>}
                        {val(d.agency)   && <span>Agency: <strong className="report-title">{d.agency}</strong></span>}
                        {val(d.startDate) && <span>Dates: <strong className="report-title">{d.startDate}{d.endDate ? ` - ${d.endDate}` : ''}</strong></span>}
                        {val(d.deliverables) && <span>Deliverable: <strong className="report-title">{d.deliverables}</strong></span>}
                      </div>
                    </div>

                    {(val(d.postImage) || val(d.videoUrl)) && (
                      <div className="report-panel-soft backdrop-blur-sm p-3 sm:p-4 rounded-lg sm:rounded-xl">
                        <h4 className="report-title font-semibold mb-2 sm:mb-3 text-sm sm:text-base md:text-lg">
                          URL Media
                        </h4>
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-3 sm:gap-4">
                          {val(d.postImage) && (
                            <div className="report-chip media-frame p-2 sm:p-3 rounded">
                              <div className="report-subtitle text-xs sm:text-sm mb-2">Post Image</div>
                              <img
                                src={proxiedMediaUrl(d.postImage)}
                                alt="Instagram post"
                                className="w-full h-48 object-cover rounded-lg border border-[var(--border-subtle)]"
                              />
                            </div>
                          )}
                          {/* {val(d.videoUrl) && (
                            <div className="report-chip media-frame p-2 sm:p-3 rounded">
                              <div className="report-subtitle text-xs sm:text-sm mb-2">Video Preview</div>
                              <video
                                src={proxiedMediaUrl(d.videoUrl)}
                                controls
                                className="w-full h-48 object-cover rounded-lg border border-[var(--border-subtle)]"
                              />
                            </div>
                          )} */}
                        </div>
                      </div>
                    )}

                    {/* Overall Metrics */}
                    <div className="report-panel-soft backdrop-blur-sm p-3 sm:p-4 rounded-lg sm:rounded-xl">
                      <h4 className="report-title font-semibold mb-2 sm:mb-3 text-sm sm:text-base md:text-lg">
                        Overall Campaign Metrics
                      </h4>
                      <div className="grid grid-cols-2 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-2 sm:gap-3 text-xs sm:text-sm">
                        {[
                          ['Views',      val(d.performance?.totalViews)  || val(d.instagram?.views)],
                          ['Likes',      val(d.performance?.totalLikes)  || val(d.instagram?.likes)],
                          ['Comments',   val(d.performance?.totalComments) || val(d.instagram?.comments)],
                          ['Shares',     val(d.performance?.totalShares) || val(d.instagram?.shares)],
                          ['Saves',      val(d.performance?.totalSaves)  || val(d.instagram?.saves)],
                          ['Reach',      val(d.performance?.totalReach)  || val(d.instagram?.reach)],
                          ['Eng. Rate',  val(d.instagram?.engagementRate)],
                          ['Engagement', val(d.performance?.totalEngagement)],
                          ['Budget',     val(formatBudget(d.financial))],
                          ['CPC',        val(d.financial?.cpc)],
                          ['CPV',        val(d.financial?.cpv)],
                          ['CPE',        val(d.financial?.cpe)],
                        ].map(([label, value]) => value ? (
                          <div key={label} className="report-chip p-2 sm:p-3 rounded">
                            <span className="report-subtitle">{label}:</span>
                            <span className="font-semibold report-title block mt-1">{value}</span>
                          </div>
                        ) : null)}
                      </div>
                      {val(d.performance?.keyLearnings) && (
                        <div className="report-chip mt-3 sm:mt-4 p-2 sm:p-3 rounded">
                          <span className="font-semibold report-title text-sm sm:text-base">Learnings:</span>
                          <p className="report-subtitle mt-1 text-xs sm:text-sm leading-relaxed">{d.performance.keyLearnings}</p>
                        </div>
                      )}
                      {[
                        ['CPC Goal',        val(d.financial?.cpcGoal)],
                        ['CPV Goal',        val(d.financial?.cpvGoal)],
                        ['CPC Calculation', val(d.financial?.cpcCalculation)],
                        ['CPV Calculation', val(d.financial?.cpvCalculation)],
                      ].some(([, value]) => value) && (
                        <div className="mt-3 sm:mt-4 grid grid-cols-1 md:grid-cols-2 gap-2 sm:gap-3 text-xs sm:text-sm">
                          {[
                            ['CPC Goal',        val(d.financial?.cpcGoal)],
                            ['CPV Goal',        val(d.financial?.cpvGoal)],
                            ['CPC Calculation', val(d.financial?.cpcCalculation)],
                            ['CPV Calculation', val(d.financial?.cpvCalculation)],
                          ].map(([label, value]) => value ? (
                            <div key={label} className="report-chip p-2 sm:p-3 rounded">
                              <span className="report-subtitle">{label}:</span>
                              <span className="font-semibold report-title block mt-1">{value}</span>
                            </div>
                          ) : null)}
                        </div>
                      )}
                    </div>

                    {/* Creator Data */}
                    {d.creators && d.creators.length > 0 && (
                      <div className="space-y-3 sm:space-y-4">
                        {d.creators.map((creator, idx) => (
                          <div key={idx} className="report-panel-soft backdrop-blur-sm p-3 sm:p-4 rounded-lg sm:rounded-xl">
                            <h4 className="report-title font-semibold mb-2 sm:mb-3 text-sm sm:text-base md:text-lg">
                              Creator: {creator.name || 'Unknown Creator'}
                            </h4>
                            <div className="grid grid-cols-2 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-2 sm:gap-3 text-xs sm:text-sm">
                              {[
                                ['Views',       val(creator.views)],
                                ['Likes',       val(creator.likes)],
                                ['Comments',    val(creator.comments)],
                                ['Shares',      val(creator.shares)],
                                ['Saves',       val(creator.saves)],
                                ['Reach',       val(creator.reach)],
                                ['Eng. Rate',   val(creator.engagementRate)],
                                ['Interactions',val(creator.interactions)],
                                ['Platform',    val(creator.platform)],
                                ['CPC',         val(creator.cpc) || val(d.financial?.cpc)],
                                ['CPV',         val(creator.cpv) || val(d.financial?.cpv)],
                              ].map(([label, value]) => value ? (
                                <div key={label} className="report-chip p-2 sm:p-3 rounded">
                                  <span className="report-subtitle">{label}:</span>
                                  <span className="font-semibold report-title block mt-1">{value}</span>
                                </div>
                              ) : null)}
                            </div>
                            {[
                              ['CPC Goal',        val(creator.cpcGoal) || val(d.financial?.cpcGoal)],
                              ['CPV Goal',        val(creator.cpvGoal) || val(d.financial?.cpvGoal)],
                              ['CPC Calculation', val(creator.cpcCalculation) || val(d.financial?.cpcCalculation)],
                              ['CPV Calculation', val(creator.cpvCalculation) || val(d.financial?.cpvCalculation)],
                            ].some(([, value]) => value) && (
                              <div className="mt-3 sm:mt-4 grid grid-cols-1 md:grid-cols-2 gap-2 sm:gap-3 text-xs sm:text-sm">
                                {[
                                  ['CPC Goal',        val(creator.cpcGoal) || val(d.financial?.cpcGoal)],
                                  ['CPV Goal',        val(creator.cpvGoal) || val(d.financial?.cpvGoal)],
                                  ['CPC Calculation', val(creator.cpcCalculation) || val(d.financial?.cpcCalculation)],
                                  ['CPV Calculation', val(creator.cpvCalculation) || val(d.financial?.cpvCalculation)],
                                ].map(([label, value]) => value ? (
                                  <div key={label} className="report-chip p-2 sm:p-3 rounded">
                                    <span className="report-subtitle">{label}:</span>
                                    <span className="font-semibold report-title block mt-1">{value}</span>
                                  </div>
                                ) : null)}
                              </div>
                            )}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                  );
                  })()}
                </div>

                {/* Action Buttons */}
                <div className="flex flex-col gap-3 sm:gap-4">
                  {/* Google Slides Link */}
                  {result.google_slides_link && (
                    <a
                      href={result.google_slides_link}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="report-link-button w-full font-semibold py-3 sm:py-4 px-4 sm:px-6 rounded-lg sm:rounded-xl transition-all duration-300 shadow-lg hover:shadow-xl text-center flex items-center justify-center gap-2 sm:gap-3 transform hover:scale-[1.01] sm:hover:scale-[1.02] active:scale-[0.99] sm:active:scale-[0.98]"
                    >
                      <svg className="w-4 h-4 sm:w-5 sm:h-5 md:w-6 md:h-6" fill="currentColor" viewBox="0 0 24 24">
                        <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                      </svg>
                      <span className="text-sm sm:text-base md:text-lg">Open in Google Slides</span>
                    </a>
                  )}

                  {/* New Report Button */}
                  <div className="flex justify-center">
                    <button
                      onClick={resetForm}
                      className="report-button-secondary font-semibold py-2.5 sm:py-3 px-6 sm:px-8 rounded-lg sm:rounded-xl transition-all duration-300 shadow-lg hover:shadow-xl backdrop-blur-sm transform hover:scale-[1.01] sm:hover:scale-[1.02] active:scale-[0.99] sm:active:scale-[0.98]"
                    >
                      <span className="text-sm sm:text-base md:text-lg">Create New Report</span>
                    </button>
                  </div>
                </div>
              </div>
            )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;
