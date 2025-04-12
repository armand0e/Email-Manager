from datetime import datetime, timedelta, timezone
import pytz
import re
import logging

# Get logger for this module
logger = logging.getLogger(__name__)

class PriorityScorer:
    def __init__(self):
        # Define sender importance with more granular scoring
        self.sender_importance = {
            'boss': 5,
            'manager': 4,
            'director': 4,
            'vp': 4,
            'team': 3,
            'client': 3,
            'customer': 3,
            'colleague': 2,
            'hr': 2,
            'support': 2,
            'noreply': 0,
            'notification': 0
        }
        
        # Define time-based weights with more granular scoring
        self.time_weights = {
            'within_15min': 5,
            'within_hour': 4,
            'within_4hours': 3,
            'within_day': 2,
            'within_week': 1
        }

        # Define urgency indicators with weights
        self.urgency_indicators = {
            'urgent': 5,
            'asap': 5,
            'immediate': 5,
            'important': 4,
            'priority': 4,
            'deadline': 4,
            'reminder': 3,
            'follow up': 3,
            'action required': 4,
            'response needed': 4,
            'please review': 3,
            'attention': 3
        }

        # Define subject patterns that indicate importance
        self.important_patterns = [
            (r're:.*', 2),  # Replies
            (r'fwd:.*', 1),  # Forwards
            (r'meeting', 3),
            (r'call', 3),
            (r'presentation', 3),
            (r'report', 2),
            (r'update', 2),
            (r'status', 2)
        ]

        # Define weights for different scoring factors
        # These could potentially be loaded from a config file
        self.weights = {
            'sender': 1.5,
            'keywords': 1.0,
            'time_sensitivity': 1.2,
            'category': 0.8,
            'entities': 1.1
        }
        # Define keywords associated with different priority levels
        # More keywords can be added or refined
        self.priority_keywords = {
            'high': ['urgent', 'important', 'action required', 'asap', 'immediately', 'critical', 'alert', 'deadline', 'final notice'],
            'medium': ['meeting', 'update', 'follow up', 'reminder', 'project', 'task', 'question', 'response needed'],
            'low': ['newsletter', 'promotion', 'update', 'summary', 'report', 'optional', 'fyi', 'announcement']
        }
        # Define important senders (example, customize as needed)
        # Could be loaded from a user config or settings
        self.important_senders = {
            'boss@example.com', 
            'client@example.com',
            'projectmanager@example.com'
        }
        logger.info("PriorityScorer initialized with default weights and keywords.")

    def score_emails(self, emails):
        """Score and prioritize emails"""
        scored_emails = []
        
        for email in emails:
            # Calculate sender score
            sender_score = self._calculate_sender_score(email['sender'])
            
            # Calculate time score
            time_score = self._calculate_time_score(email['date'])
            
            # Calculate content score
            content_score = self._calculate_content_score(email['subject'], email['snippet'])
            
            # Calculate pattern score
            pattern_score = self._calculate_pattern_score(email['subject'])
            
            # Combine scores with weighted average
            total_score = (
                sender_score * 0.25 +
                time_score * 0.25 +
                content_score * 0.25 +
                pattern_score * 0.25
            )
            
            # Apply bonus for recent replies
            if self._is_recent_reply(email):
                total_score += 2
            
            # Cap the total score at 10
            total_score = min(total_score, 10)
            
            # Determine priority level
            priority = self._determine_priority(total_score)
            
            # Add scoring results to email
            email['score'] = total_score
            email['priority'] = priority
            email['sender_score'] = sender_score
            email['time_score'] = time_score
            email['content_score'] = content_score
            email['pattern_score'] = pattern_score
            
            scored_emails.append(email)
        
        # Sort emails by total score (descending)
        return sorted(scored_emails, key=lambda x: x['score'], reverse=True)

    def _calculate_sender_score(self, sender):
        """Calculate sender importance score"""
        sender_lower = sender.lower()
        score = 0
        
        # Check for exact matches
        for keyword, weight in self.sender_importance.items():
            if keyword in sender_lower:
                score = max(score, weight)
        
        # Check for domain importance
        if '@' in sender_lower:
            domain = sender_lower.split('@')[1]
            if 'company.com' in domain:  # Replace with actual company domain
                score = max(score, 3)
        
        return score

    def _calculate_time_score(self, date_str):
        """Calculate time-based score"""
        try:
            # Try parsing with timezone
            email_date = datetime.strptime(date_str, '%a, %d %b %Y %H:%M:%S %z')
        except ValueError:
            try:
                # Try parsing without timezone
                email_date = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
                # Make it timezone-aware using UTC
                email_date = email_date.replace(tzinfo=pytz.UTC)
            except ValueError:
                return 0
        
        # Get current time in UTC
        now = datetime.now(pytz.UTC)
        time_diff = now - email_date
        
        if time_diff < timedelta(minutes=15):
            return self.time_weights['within_15min']
        elif time_diff < timedelta(hours=1):
            return self.time_weights['within_hour']
        elif time_diff < timedelta(hours=4):
            return self.time_weights['within_4hours']
        elif time_diff < timedelta(days=1):
            return self.time_weights['within_day']
        elif time_diff < timedelta(weeks=1):
            return self.time_weights['within_week']
        
        return 0

    def _calculate_content_score(self, subject, snippet):
        """Calculate score based on email content"""
        score = 0
        text = f"{subject} {snippet}".lower()
        
        # Check for urgency indicators
        for indicator, weight in self.urgency_indicators.items():
            if indicator in text:
                score = max(score, weight)
        
        # Check for question marks (indicates response needed)
        if '?' in text:
            score += 1
        
        # Check for action words
        action_words = ['please', 'need', 'request', 'required', 'should', 'must']
        for word in action_words:
            if word in text:
                score += 1
        
        return min(score, 5)  # Cap at 5

    def _calculate_pattern_score(self, subject):
        """Calculate score based on subject patterns"""
        score = 0
        subject_lower = subject.lower()
        
        for pattern, weight in self.important_patterns:
            if re.match(pattern, subject_lower):
                score += weight
        
        return min(score, 5)  # Cap at 5

    def _is_recent_reply(self, email):
        """Check if email is a recent reply"""
        try:
            email_date = datetime.strptime(email['date'], '%a, %d %b %Y %H:%M:%S %z')
        except ValueError:
            try:
                email_date = datetime.strptime(email['date'], '%Y-%m-%d %H:%M:%S')
                email_date = email_date.replace(tzinfo=pytz.UTC)
            except ValueError:
                return False
        
        now = datetime.now(pytz.UTC)
        return (now - email_date < timedelta(hours=24) and 
                email['subject'].lower().startswith('re:'))

    def _determine_priority(self, score):
        """Determine priority level based on total score"""
        if score >= 8:
            return 'High'
        elif score >= 5:
            return 'Medium'
        else:
            return 'Low'

    def _score_sender(self, sender):
        """Score based on whether the sender is in the important list."""
        if not sender:
             return 0
        # Extract email address if sender format is "Name <email@example.com>"
        match = re.search(r'<([^>]+)>', sender)
        email_address = match.group(1).lower() if match else sender.lower()
        
        if email_address in self.important_senders:
             logger.debug(f"High sender score for: {email_address}")
             return 1.0 # Max score for important sender
        else:
             return 0.2 # Base score for unknown/less important sender

    def _score_keywords(self, tokens):
        """Score based on the presence and frequency of priority keywords."""
        if not tokens:
            return 0
        score = 0
        text = " ".join(tokens) # Join tokens for easier regex matching on phrases
        
        # Check for high priority keywords/phrases
        for keyword in self.priority_keywords['high']:
            if keyword in text:
                score += 0.8 # Significant boost for high priority
                logger.debug(f"High priority keyword found: '{keyword}'")
                break # Stop after first high priority match for simplicity
        
        # Check for medium priority keywords if no high priority found yet
        if score == 0:
             for keyword in self.priority_keywords['medium']:
                  if keyword in text:
                       score += 0.4 # Medium boost
                       logger.debug(f"Medium priority keyword found: '{keyword}'")
                       break

        # Penalize slightly for low priority keywords (optional)
        # for keyword in self.priority_keywords['low']:
        #      if keyword in text:
        #           score -= 0.1
        #           logger.debug(f"Low priority keyword found: '{keyword}', reducing score slightly.")
        #           break 

        # Normalize score (crude normalization)
        return max(0, min(score, 1.0))

    def _score_time_sensitivity(self, date_str):
        """Score based on how recent the email is."""
        if not date_str:
            return 0.1 # Default low score if no date
        try:
            # Attempt to parse the date string (assuming format from email_client)
            # This needs to be robust to the actual format provided
            # Re-use the parsing logic from app.py or ensure consistency
            # For now, using a simplified approach assuming common formats
            date_obj = None
            formats_to_try = [
                 '%a, %d %b %Y %H:%M:%S %z',
                 '%Y-%m-%d %H:%M:%S%z', # Example format from exchangelib string
                 '%Y-%m-%d %H:%M:%S'
            ]
            for fmt in formats_to_try:
                 try:
                      # Remove potential fractional seconds if present
                      date_str_cleaned = str(date_str).split('.')[0]
                      date_obj = datetime.strptime(date_str_cleaned, fmt)
                      # Handle potential timezone offset string like +00:00
                      if date_obj.tzinfo is None and '%z' in fmt:
                           # Manually parse offset if strptime didn't handle it
                           tz_match = re.search(r'([+-]\d{2}):?(\d{2})$', str(date_str))
                           if tz_match:
                                hours, minutes = map(int, tz_match.groups())
                                sign = -1 if tz_match.group(1).startswith('-') else 1
                                offset = timedelta(hours=sign*hours, minutes=sign*minutes)
                                date_obj = date_obj.replace(tzinfo=timezone(offset))
                      break
                 except ValueError:
                      continue
            
            if not date_obj:
                logger.warning(f"Could not parse date for time sensitivity: {date_str}")
                return 0.1

            # Ensure date_obj is timezone-aware for comparison
            if date_obj.tzinfo is None or date_obj.tzinfo.utcoffset(date_obj) is None:
                 # If naive, assume UTC or local? For consistency, let's assume UTC
                 date_obj = date_obj.replace(tzinfo=timezone.utc)
                 logger.debug(f"Assumed UTC for naive date: {date_str}")

            now = datetime.now(timezone.utc) # Compare against UTC now
            age = now - date_obj

            # Scoring based on age
            if age < timedelta(hours=1):
                logger.debug(f"High time score: email age {age}")
                return 1.0 # Very recent
            elif age < timedelta(hours=6):
                return 0.8 # Recent
            elif age < timedelta(days=1):
                return 0.6 # Within a day
            elif age < timedelta(days=3):
                return 0.4 # Few days old
            else:
                return 0.1 # Older

        except Exception as e:
            logger.error(f"Error calculating time sensitivity for date '{date_str}': {e}", exc_info=True)
            return 0.1 # Default low score on error

    def _score_category(self, category):
        """Assign score based on detected category."""
        # Higher score for potentially more important categories
        category_scores = {
            "Work": 0.8,
            "Personal": 0.7,
            "Notifications": 0.5, # Can be important but often informational
            "Other": 0.3,
            "Newsletters": 0.1 # Generally lowest priority
        }
        score = category_scores.get(category, 0.3) # Default to 'Other' score
        logger.debug(f"Category score for '{category}': {score}")
        return score

    def _score_entities(self, entities):
        """Score based on the presence of important entity types (e.g., ORG, PERSON)."""
        if not entities:
            return 0
        score = 0
        # Boost score if specific organizations or people are mentioned (examples)
        important_entities = {'ORG', 'PERSON'} 
        for _, label in entities:
             if label in important_entities:
                  score += 0.3 # Add score for each relevant entity type found
                  logger.debug(f"Entity type '{label}' found, increasing score.")
        
        return min(score, 1.0) # Cap score at 1.0

    def _calculate_final_score(self, scores):
        """Calculate weighted average score and map to priority level."""
        final_score = 0
        total_weight = 0
        for factor, score in scores.items():
            weight = self.weights.get(factor, 1.0) # Default weight is 1
            final_score += score * weight
            total_weight += weight
            logger.debug(f"Score factor '{factor}': score={score:.2f}, weight={weight:.2f}")
        
        if total_weight == 0: return 0 # Avoid division by zero
        
        normalized_score = final_score / total_weight
        logger.info(f"Calculated normalized score: {normalized_score:.3f}")
        return normalized_score

    def _map_score_to_priority(self, score):
        """Map numerical score to High, Medium, Low priority."""
        if score >= 0.7:
            return "High"
        elif score >= 0.4:
            return "Medium"
        else:
            return "Low" 