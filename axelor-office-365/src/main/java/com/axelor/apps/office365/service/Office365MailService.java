/*
 * Axelor Business Solutions
 *
 * Copyright (C) 2020 Axelor (<http://axelor.com>).
 *
 * This program is free software: you can redistribute it and/or  modify
 * it under the terms of the GNU Affero General Public License, version 3,
 * as published by the Free Software Foundation.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
package com.axelor.apps.office365.service;

import com.axelor.apps.base.db.Partner;
import com.axelor.apps.base.db.repo.PartnerRepository;
import com.axelor.apps.message.db.EmailAccount;
import com.axelor.apps.message.db.EmailAddress;
import com.axelor.apps.message.db.Message;
import com.axelor.apps.message.db.repo.EmailAccountRepository;
import com.axelor.apps.message.db.repo.EmailAddressRepository;
import com.axelor.apps.message.db.repo.MessageRepository;
import com.axelor.apps.office.db.Office365Account;
import com.axelor.apps.office.db.Office365Message;
import com.axelor.apps.office.db.repo.Office365MessageRepository;
import com.axelor.auth.db.User;
import com.axelor.auth.db.repo.UserRepository;
import com.axelor.common.ObjectUtils;
import com.axelor.common.StringUtils;
import com.axelor.exception.service.TraceBackService;
import com.google.inject.Inject;
import com.google.inject.persist.Transactional;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.HashSet;
import java.util.Optional;
import java.util.Set;
import wslite.json.JSONArray;
import wslite.json.JSONException;
import wslite.json.JSONObject;

public class Office365MailService {

  @Inject private EmailAddressRepository emailAddressRepo;
  @Inject private EmailAccountRepository emailAccountRepo;
  @Inject private MessageRepository messageRepo;
  @Inject private Office365MessageRepository office365MessageRepo;
  @Inject private UserRepository userRepo;
  @Inject private PartnerRepository partnerRepository;

  @Inject private Office365Service office365Service;

  @Transactional
  @SuppressWarnings("unchecked")
  public void createMessage(
      JSONObject jsonObject, Office365Account office365Account, LocalDateTime lastSyncOn) {

    if (jsonObject == null) {
      return;
    }

    try {
      String officeMessageId = office365Service.processJsonValue("id", jsonObject);
      Office365Message office365Message =
          office365MessageRepo.findOffice365Message(officeMessageId, office365Account);

      Message message = office365Message == null ? null : office365Message.getMessage();
      if (message == null) {
        message = new Message();
        message.setIsOffice365Object(true);

      } else if (!office365Service.needUpdation(
          jsonObject, lastSyncOn, message.getCreatedOn(), message.getUpdatedOn())) {
        return;
      }

      message.setSubject(office365Service.processJsonValue("subject", jsonObject));
      message.setMediaTypeSelect(MessageRepository.MEDIA_TYPE_EMAIL);
      message.setSentDateT(
          office365Service.processLocalDateTimeValue(
              jsonObject, "sentDateTime", ZoneId.systemDefault()));
      message.setReceivedDateT(
          office365Service.processLocalDateTimeValue(
              jsonObject, "receivedDateTime", ZoneId.systemDefault()));

      if (jsonObject.containsKey("body")) {
        JSONObject bodyJsonObj = (JSONObject) jsonObject.get("body");
        if (bodyJsonObj.containsKey("content")) {
          message.setContent(office365Service.processJsonValue("content", bodyJsonObj));
        } else {
          message.setContent(office365Service.processJsonValue("bodyPreview", jsonObject));
        }
      }

      if (jsonObject.getBoolean("isDraft")) {
        message.setStatusSelect(MessageRepository.STATUS_DRAFT);
      } else if (jsonObject.getBoolean("isRead")) {
        message.setTypeSelect(MessageRepository.TYPE_RECEIVED);
      } else {
        message.setStatusSelect(MessageRepository.STATUS_SENT);
        message.setTypeSelect(MessageRepository.TYPE_SENT);
      }

      JSONObject fromJsonObj = (JSONObject) jsonObject.getOrDefault("from", JSONObject.NULL);
      message.setFromEmailAddress(getEmailAddress(fromJsonObj));

      message.setToEmailAddressSet(getEmailAddressSet(jsonObject, "toRecipients"));
      message.setReplyToEmailAddressSet(getEmailAddressSet(jsonObject, "replyTo"));
      message.setCcEmailAddressSet(getEmailAddressSet(jsonObject, "ccRecipients"));
      message.setBccEmailAddressSet(getEmailAddressSet(jsonObject, "bccRecipients"));
      setSender(jsonObject, message);

      messageRepo.save(message);

      if (office365Message == null) {
        office365Message = new Office365Message();
        office365Message.setOffice365Id(officeMessageId);
        office365Message.setOffice365Account(office365Account);
        message.addOffice365MessageListItem(office365Message);
        office365MessageRepo.save(office365Message);
      }

    } catch (Exception e) {
      TraceBackService.trace(e);
    }
  }

  @Transactional
  public void createOffice365Mail(
      Message message, Office365Account office365Account, String accessToken) {

    try {
      JSONObject messageJsonObject = setOffice365MailValues(message);
      String office365Id = null;
      Office365Message office365Messge = null;
      if (!ObjectUtils.isEmpty(message.getOffice365MessageList())) {
        Optional<Office365Message> office365MessageOpt =
            message.getOffice365MessageList().stream()
                .filter(oPartner -> oPartner.getOffice365Account().equals(office365Account))
                .findFirst();
        if (office365MessageOpt.isPresent()) {
          office365Messge = office365MessageOpt.get();
          office365Id = office365Messge.getOffice365Id();
        }
      }

      if (office365Messge == null) {
        office365Messge = new Office365Message();
        office365Messge.setOffice365Account(office365Account);
        message.addOffice365MessageListItem(office365Messge);
      }

      office365Id =
          office365Service.createOffice365Object(
              Office365Service.MAIL_URL, messageJsonObject, accessToken, office365Id, "Messages");

      office365Messge.setOffice365Id(office365Id);
      office365MessageRepo.save(office365Messge);

    } catch (Exception e) {
      TraceBackService.trace(e);
    }
  }

  private JSONObject setOffice365MailValues(Message message) throws JSONException {

    JSONObject messageJsonObject = new JSONObject();
    office365Service.putObjValue(messageJsonObject, "subject", message.getSubject());

    JSONObject bodyJsonObject = new JSONObject();
    bodyJsonObject.put("content", message.getContent());
    bodyJsonObject.put("contentType", "html");
    messageJsonObject.put("body", (Object) bodyJsonObject);
    messageJsonObject.put("importance", "Low");

    if (MessageRepository.STATUS_DRAFT == message.getStatusSelect()) {
      messageJsonObject.put("isDraft", true);
    }

    if (message.getSentDateT() != null) {
      LocalDateTime sentDateT =
          LocalDateTime.ofInstant(
              Instant.parse(message.getSentDateT().toString() + "Z"), ZoneId.of("UTC"));
      messageJsonObject.put("sentDateTime", sentDateT.toString() + "Z");
    }

    manageOffice365EmailAddresses(
        messageJsonObject, "bccRecipients", message.getBccEmailAddressSet());
    manageOffice365EmailAddresses(
        messageJsonObject, "ccRecipients", message.getCcEmailAddressSet());
    manageOffice365EmailAddresses(
        messageJsonObject, "replyTo", message.getReplyToEmailAddressSet());
    manageOffice365EmailAddresses(
        messageJsonObject, "toRecipients", message.getToEmailAddressSet());

    office365Service.putUserEmailAddress(message.getSenderUser(), messageJsonObject, "sender");
    JSONObject fromJsonObj = createOffice365EmailAddress(message.getFromEmailAddress());
    if (fromJsonObj != null) {
      messageJsonObject.put("from", (Object) fromJsonObj);
    }

    return messageJsonObject;
  }

  private EmailAddress getEmailAddress(JSONObject jsonObject) throws JSONException {

    EmailAddress emailAddress = null;
    if (jsonObject != JSONObject.NULL) {
      JSONObject emailAddJsonObj = jsonObject.getJSONObject("emailAddress");

      String emailAddressStr = emailAddJsonObj.getString("address");
      if (!StringUtils.isBlank(emailAddressStr)) {
        emailAddress = emailAddressRepo.findByAddress(emailAddressStr);
      }

      if (emailAddress == null) {
        emailAddress = new EmailAddress();
        emailAddress.setAddress(emailAddressStr);
      }
      emailAddress.setName(emailAddJsonObj.getString("name"));
    }

    return emailAddress;
  }

  private Set<EmailAddress> getEmailAddressSet(JSONObject jsonObject, String key)
      throws JSONException {

    Set<EmailAddress> toEmailAddressSet = null;
    JSONArray toJsonArr = jsonObject.getJSONArray(key);
    if (toJsonArr != null && toJsonArr.size() > 0) {
      toEmailAddressSet = new HashSet<>();
      for (Object obj : toJsonArr) {
        JSONObject toJsonObj = (JSONObject) obj;
        toEmailAddressSet.add(getEmailAddress(toJsonObj));
      }
    }

    return toEmailAddressSet;
  }

  private void setSender(JSONObject jsonObject, Message message) throws JSONException {

    JSONObject senderJsonObj = jsonObject.getJSONObject("sender");
    if (senderJsonObj != JSONObject.NULL) {

      JSONObject emailAddressJsonObj = senderJsonObj.getJSONObject("emailAddress");
      String email = emailAddressJsonObj.getString("address");
      String name = emailAddressJsonObj.getString("name");

      User user = userRepo.findByCode(name);
      if (user == null) {
        user = new User();
        user.setEmail(email);
        user.setName(name);
        user.setCode(name);
        user.setPassword(name);

        EmailAddress emailAddress = getEmailAddress(senderJsonObj);
        Partner partner =
            partnerRepository
                .all()
                .filter("self.emailAddress = :emailAddress")
                .bind("emailAddress", emailAddress)
                .fetchOne();
        user.setPartner(partner);
      }
      message.setSenderUser(user);

      EmailAccount emailAccount = emailAccountRepo.findByName(name);
      if (emailAccount == null) {
        emailAccount = new EmailAccount();
        emailAccount.setServerTypeSelect(EmailAccountRepository.SERVER_TYPE_SMTP);
        emailAccount.setSecuritySelect(EmailAccountRepository.SECURITY_NONE);
        emailAccount.setName(name);
        emailAccount.setLogin(email);
        emailAccount.setHost("Microsoft");
      }
      message.setMailAccount(emailAccount);
    }
  }

  private void manageOffice365EmailAddresses(
      JSONObject jsonObject, String key, Set<EmailAddress> emailAddressSet) throws JSONException {

    if (ObjectUtils.isEmpty(emailAddressSet)) {
      return;
    }

    JSONArray emailJsonArr = new JSONArray();
    for (EmailAddress emailAddress : emailAddressSet) {
      JSONObject emailJsonObject = createOffice365EmailAddress(emailAddress);
      if (emailJsonObject == null) {
        continue;
      }

      emailJsonArr.add(emailJsonObject);
    }
    jsonObject.put(key, (Object) emailJsonArr);
  }

  private JSONObject createOffice365EmailAddress(EmailAddress emailAddress) throws JSONException {

    if (emailAddress == null || StringUtils.isBlank(emailAddress.getAddress())) {
      return null;
    }

    JSONObject emailJsonObj = new JSONObject();
    office365Service.putObjValue(emailJsonObj, "address", emailAddress.getAddress());
    office365Service.putObjValue(emailJsonObj, "name", emailAddress.getName());

    JSONObject emailAddressObj = new JSONObject();
    emailAddressObj.put("emailAddress", (Object) emailJsonObj);

    return emailAddressObj;
  }
}
