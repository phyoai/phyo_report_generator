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
  const [contentUrls, setContentUrls] = useState(['', '', '']);
  const [budgetInrValues, setBudgetInrValues] = useState(['', '', '']);
  const [images, setImages] = useState([]);
  const [previews, setPreviews] = useState([]);
  const [loading, setLoading] = useState(false);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [templatePreview, setTemplatePreview] = useState(null);
  const [generatedPreview, setGeneratedPreview] = useState(null);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);

  useEffect(() => {
    return () => {
      previews.forEach((url) => URL.revokeObjectURL(url));
    };
  }, [previews]);

  const loadTemplatePreview = async () => {
    setPreviewLoading(true);

    try {
      const res = await fetch(`${apiUrl}/api/template-preview`);
      const data = await res.json();

      if (data.success) {
        setTemplatePreview(data);
      }
    } catch (err) {
      console.error('Template preview failed:', err);
    } finally {
      setPreviewLoading(false);
    }
  };

  useEffect(() => {
    loadTemplatePreview();
  }, []);

  const loadGeneratedPreview = async (filename) => {
    if (!filename) {
      setGeneratedPreview(null);
      return;
    }

    setPreviewLoading(true);

    try {
      const res = await fetch(`${apiUrl}/api/report-preview/${encodeURIComponent(filename)}`);
      const data = await res.json();

      if (data.success) {
        setGeneratedPreview(data);
      }
    } catch (err) {
      console.error('Generated report preview failed:', err);
    } finally {
      setPreviewLoading(false);
    }
  };

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
    
    const combinedUrls = contentUrls.map((url) => url.trim()).filter(Boolean).join('\n');
    const submittedUrls = contentUrls.map((url) => url.trim()).filter(Boolean);

    if (!prompt.trim() && images.length === 0 && !combinedUrls) {
      setError('Please provide a prompt, images, or a post URL');
      return;
    }

    setLoading(true);
    setError(null);
    setResult(null);

    try {
      const formData = new FormData();
      formData.append('prompt', prompt);
      formData.append(
        'budget_inr_values',
        JSON.stringify(budgetInrValues.map((value) => value.trim()))
      );
      if (combinedUrls) {
        formData.append('instagram_post_url', combinedUrls);
      }

      console.log('[URL SUBMIT] Content URL slots:', contentUrls);
      console.log('[URL SUBMIT] Final submitted URLs:', submittedUrls);
      console.log('[URL SUBMIT] Final combined payload:', combinedUrls);

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
        await loadGeneratedPreview(data.filename);
      } else {
        setError(data.error || 'Failed to generate report');
      }
    } catch (err) {
      setError('Error: ' + err.message);
    } finally {
      setLoading(false);
    }
  };

  const resetForm = () => {
    setPrompt('');
    setContentUrls(['', '', '']);
    setBudgetInrValues(['', '', '']);
    setImages([]);
    setPreviews([]);
    setGeneratedPreview(null);
    setResult(null);
    setError(null);
  };

  const activePreview = generatedPreview || templatePreview;
  const previewTitle = generatedPreview ? 'Generated PPT Preview' : 'PPT Template Preview';
  const previewHeadline = generatedPreview ? 'Latest generated deck' : 'Presentation template library';
  const previewNote = generatedPreview
    ? 'This preview reflects the latest generated PPTX report with refreshed slide visuals.'
    : 'Edit `backend/Campaign_Report_PyroMedia.pptx`, save it, then refresh this preview.';
  const refreshPreview = generatedPreview?.slides?.length && result?.filename
    ? () => loadGeneratedPreview(result.filename)
    : loadTemplatePreview;

  return (
    <div className="report-app min-h-screen w-full relative overflow-x-hidden">
      <div className="report-pattern fixed inset-0 w-full h-full opacity-50 pointer-events-none">
        <div className="absolute inset-0 w-full h-full transform rotate-45"></div>
        <div className="absolute inset-0 w-full h-full"></div>
      </div>
      
      <div className="w-full min-h-screen flex flex-col relative z-10">
        <div className="container mx-auto px-3 sm:px-4 py-3 sm:py-5 md:py-6 flex-1">
          <div className="max-w-5xl mx-auto">
            <div className="report-shell backdrop-blur-xl rounded-2xl sm:rounded-3xl p-4 sm:p-6 md:p-8 lg:p-10">
            <div className="report-hero report-hero-compact mb-4 sm:mb-5 md:mb-6">
              <div className="report-hero-logo-shell">
                <img
                  src="/pyromedia-logo.png"
                  alt="PyroMedia logo"
                  className="report-hero-logo"
                />
              </div>
              <h5 className="report-hero-subhead">Report Generator</h5>
            </div>

            <form onSubmit={handleSubmit} className="space-y-4 sm:space-y-5 md:space-y-6">
              
              <div className="report-form-grid">
                <div className="report-section report-form-grid-main space-y-2">
                  <div className="report-section-title">Content Source</div>
                  <label className="report-label block text-sm sm:text-base font-semibold mb-1">
                    Post / Reel / Video URLs
                  </label>
                  <div className="report-url-grid">
                    {contentUrls.map((url, index) => (
                      <div key={`content-url-${index}`} className="report-dropzone report-url-slot flex items-center rounded-lg sm:rounded-xl overflow-hidden transition-all duration-200">
                        <span className="pl-3 report-muted select-none report-url-slot-index">
                          {index + 1}
                        </span>
                        <input
                          type="text"
                          value={url}
                          onChange={(e) => {
                            const nextUrls = [...contentUrls];
                            nextUrls[index] = e.target.value;
                            setContentUrls(nextUrls);
                          }}
                          placeholder={`Paste URL ${index + 1}`}
                          disabled={loading}
                          className="flex-1 px-3 py-2 sm:py-3 bg-transparent report-title text-sm sm:text-base focus:outline-none disabled:opacity-50 disabled:cursor-not-allowed"
                        />
                      </div>
                    ))}
                  </div>
                  <p className="report-faint text-xs">Paste up to 3 links to combine them into one report. Instagram metrics will be merged across all valid URLs that return data.</p>
                </div>

                <div className="report-section report-form-grid-side space-y-2">
                  <div className="report-section-title">Campaign Budget</div>
                  <div className="report-budget-grid">
                    {budgetInrValues.map((budgetValue, index) => (
                      <div key={`budget-slot-${index}`} className="report-budget-row report-budget-row-compact">
                        <span className="report-budget-prefix">URL {index + 1}</span>
                        <input
                          type="number"
                          min="0"
                          step="0.01"
                          value={budgetValue}
                          onChange={(e) => {
                            const nextBudgets = [...budgetInrValues];
                            nextBudgets[index] = e.target.value;
                            setBudgetInrValues(nextBudgets);
                          }}
                          placeholder={contentUrls[index]?.trim() ? "Enter budget" : "Add URL first"}
                          disabled={loading || !contentUrls[index]?.trim()}
                          className="report-input report-budget-input w-full px-3 py-2 sm:py-3 rounded-lg sm:rounded-xl text-sm sm:text-base focus:outline-none disabled:opacity-50 disabled:cursor-not-allowed"
                        />
                      </div>
                    ))}
                  </div>
                  <p className="report-faint text-xs">Add only the budgets you want to use for URL-wise calculations.</p>
                </div>
              </div>
              
              
              <button
                type="submit"
                disabled={loading || (!prompt.trim() && images.length === 0 && !contentUrls.some((url) => url.trim()) && !budgetInrValues.some((value) => value.trim()))}
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
            {error && (
              <div className="report-alert-danger mt-4 sm:mt-6 backdrop-blur-md p-3 sm:p-4 rounded-lg sm:rounded-xl">
                <div className="flex flex-col sm:flex-row sm:items-center gap-1 sm:gap-0">
                  <span className="font-semibold text-sm sm:text-base">Error: </span>
                  <span className="sm:ml-2 text-sm sm:text-base break-words">{error}</span>
                </div>
              </div>
            )}
            {result && (
              <div className="mt-4 sm:mt-6 space-y-3 sm:space-y-4 md:space-y-6">

                
                <div className="report-panel backdrop-blur-md p-3 sm:p-4 md:p-6 rounded-lg sm:rounded-xl">
                  <h2 className="report-title text-lg sm:text-xl md:text-2xl font-bold mb-3 sm:mb-4 md:mb-6">
                    Extracted Data
                  </h2>
                  {(() => {
                    const d = result.extracted_data;
                    const val = (v) => (!v || v === '0' || v === 0) ? '' : v;

                    return (
                  <div className="space-y-3 sm:space-y-4 md:space-y-6">
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
                      {d.creators && d.creators.length > 1 && (
                        <div className="report-chip mt-3 sm:mt-4 p-2 sm:p-3 rounded">
                          <span className="font-semibold report-title text-sm sm:text-base">Creators Included:</span>
                          <div className="report-creator-grid mt-2">
                            {d.creators.map((creatorItem, idx) => (
                              <div key={`${creatorItem.name || 'creator'}-${idx}`} className="report-creator-card">
                                <div className="report-title font-semibold text-sm">{creatorItem.name || `Creator ${idx + 1}`}</div>
                                <div className="report-subtitle text-xs mt-1">
                                  {[
                                    creatorItem.views ? `Views ${creatorItem.views}` : '',
                                    creatorItem.likes ? `Likes ${creatorItem.likes}` : '',
                                    creatorItem.comments ? `Comments ${creatorItem.comments}` : '',
                                    creatorItem.reach ? `Reach ${creatorItem.reach}` : '',
                                  ].filter(Boolean).join(' | ') || 'Metadata captured'}
                                </div>
                              </div>
                            ))}
                          </div>
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
                <div className="flex flex-col gap-3 sm:gap-4">
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
