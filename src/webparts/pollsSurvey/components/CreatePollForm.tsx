import * as React from 'react';
import styles from './PollsSurvey.module.scss';

export interface ICreatePollFormProps {
  /** Whether form is visible */
  isOpen: boolean;
  /** Whether submission is in progress */
  isSubmitting: boolean;
  /** Callback to create a new poll */
  onCreate: (question: string, options: string[], expiryDate: string | null) => void;
  /** Callback to close the form */
  onClose: () => void;
}

/**
 * Form for admins to create a new poll.
 * - Question text field
 * - 2-6 option fields (dynamic add/remove)
 * - Optional expiry date picker
 * - Validates all fields before submission
 */
const CreatePollForm: React.FC<ICreatePollFormProps> = ({
  isOpen,
  isSubmitting,
  onCreate,
  onClose
}) => {
  const [question, setQuestion] = React.useState('');
  const [options, setOptions] = React.useState<string[]>(['', '']);
  const [expiryDate, setExpiryDate] = React.useState('');
  const [error, setError] = React.useState('');

  // Reset form when opened
  React.useEffect(() => {
    if (isOpen) {
      setQuestion('');
      setOptions(['', '']);
      setExpiryDate('');
      setError('');
    }
  }, [isOpen]);

  const handleAddOption = (): void => {
    if (options.length < 6) {
      setOptions([...options, '']);
    }
  };

  const handleRemoveOption = (index: number): void => {
    if (options.length > 2) {
      const newOptions = options.filter((_, i) => i !== index);
      setOptions(newOptions);
    }
  };

  const handleOptionChange = (index: number, value: string): void => {
    const newOptions = [...options];
    newOptions[index] = value;
    setOptions(newOptions);
  };

  const handleSubmit = (): void => {
    // Validate question
    if (!question.trim()) {
      setError('Please enter a question.');
      return;
    }

    // Validate options: at least 2 non-empty options
    const validOptions = options.map(o => o.trim()).filter(o => o.length > 0);
    if (validOptions.length < 2) {
      setError('Please provide at least 2 options.');
      return;
    }

    // Check for duplicate options
    const uniqueOptions = new Set(validOptions.map(o => o.toLowerCase()));
    if (uniqueOptions.size !== validOptions.length) {
      setError('Options must be unique.');
      return;
    }

    // Validate expiry date if provided
    if (expiryDate) {
      const expiry = new Date(expiryDate);
      if (expiry <= new Date()) {
        setError('Expiry date must be in the future.');
        return;
      }
    }

    setError('');
    onCreate(
      question.trim(),
      validOptions,
      expiryDate ? new Date(expiryDate).toISOString() : null
    );
  };

  if (!isOpen) return null;

  return (
    <div className={styles.formOverlay}>
      <div className={styles.formCard}>
        <div className={styles.formHeader}>
          <h3>Create New Poll</h3>
          <button className={styles.closeButton} onClick={onClose} title="Close">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
              <path d="M5 5L15 15M15 5L5 15" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
            </svg>
          </button>
        </div>

        <div className={styles.formBody}>
          {/* Question */}
          <div className={styles.formGroup}>
            <label className={styles.formLabel}>Question *</label>
            <input
              type="text"
              className={styles.formInput}
              placeholder="What would you like to ask?"
              value={question}
              onChange={(e) => setQuestion(e.target.value)}
              maxLength={255}
            />
          </div>

          {/* Options */}
          <div className={styles.formGroup}>
            <label className={styles.formLabel}>
              Options * <span className={styles.formHint}>({options.length}/6)</span>
            </label>
            {options.map((option, index) => (
              <div key={index} className={styles.optionRow}>
                <input
                  type="text"
                  className={styles.formInput}
                  placeholder={`Option ${index + 1}`}
                  value={option}
                  onChange={(e) => handleOptionChange(index, e.target.value)}
                  maxLength={255}
                />
                {options.length > 2 && (
                  <button
                    className={styles.removeOptionButton}
                    onClick={() => handleRemoveOption(index)}
                    title="Remove option"
                  >
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                      <path d="M4 8H12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
                    </svg>
                  </button>
                )}
              </div>
            ))}
            {options.length < 6 && (
              <button className={styles.addOptionButton} onClick={handleAddOption}>
                <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                  <path d="M8 4V12M4 8H12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
                </svg>
                Add Option
              </button>
            )}
          </div>

          {/* Expiry Date */}
          <div className={styles.formGroup}>
            <label className={styles.formLabel}>
              Expiry Date <span className={styles.formHint}>(optional)</span>
            </label>
            <input
              type="datetime-local"
              className={styles.formInput}
              value={expiryDate}
              onChange={(e) => setExpiryDate(e.target.value)}
            />
          </div>

          {/* Error message */}
          {error && <div className={styles.formError}>{error}</div>}
        </div>

        <div className={styles.formFooter}>
          <button className={styles.secondaryButton} onClick={onClose} disabled={isSubmitting}>
            Cancel
          </button>
          <button className={styles.primaryButton} onClick={handleSubmit} disabled={isSubmitting}>
            {isSubmitting ? (
              <>
                <div className={styles.spinnerSmall} />
                Creating...
              </>
            ) : (
              'Create Poll'
            )}
          </button>
        </div>
      </div>
    </div>
  );
};

export default CreatePollForm;
