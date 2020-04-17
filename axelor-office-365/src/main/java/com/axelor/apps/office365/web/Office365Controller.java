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
package com.axelor.apps.office365.web;

import com.axelor.apps.base.db.AppOffice365;
import com.axelor.apps.base.db.ICalendar;
import com.axelor.apps.base.db.Partner;
import com.axelor.apps.base.db.repo.AppOffice365Repository;
import com.axelor.apps.message.db.Message;
import com.axelor.apps.office365.service.Office365Service;
import com.axelor.apps.office365.translation.ITranslation;
import com.axelor.i18n.I18n;
import com.axelor.inject.Beans;
import com.axelor.meta.schema.actions.ActionView;
import com.axelor.rpc.ActionRequest;
import com.axelor.rpc.ActionResponse;
import com.github.scribejava.apis.MicrosoftAzureActiveDirectory20Api;
import com.github.scribejava.core.builder.ServiceBuilder;
import com.github.scribejava.core.oauth.OAuth20Service;
import com.google.inject.Inject;
import java.util.HashMap;
import java.util.Map;

public class Office365Controller {

  @Inject private Office365Service office365Service;

  public void generateUrl(ActionRequest request, ActionResponse response) throws Exception {

    AppOffice365 appOffice365 = request.getContext().asType(AppOffice365.class);

    Map<String, String> additionalParams = new HashMap<>();
    additionalParams.put("access_type", "offline");
    additionalParams.put("prompt", "consent");

    OAuth20Service authService =
        new ServiceBuilder(appOffice365.getClientId())
            .apiSecret(appOffice365.getClientSecret())
            .callback(appOffice365.getRedirectUri())
            .defaultScope(Office365Service.SCOPE)
            .build(MicrosoftAzureActiveDirectory20Api.instance());
    String authenticationUrl =
        authService
            .createAuthorizationUrlBuilder()
            .state("")
            .additionalParams(additionalParams)
            .build();
    authService.close();
    String url =
        String.format(
            "<a href='%s'>%s</a> ",
            authenticationUrl.replace("&", "&amp;"),
            I18n.get(ITranslation.OFFICE365_AUTHETICATE_TITLE));

    response.setValue("authenticationUrl", url);
    response.setAttr("authenticationUrl", "hidden", false);
  }

  public void syncContact(ActionRequest request, ActionResponse response) throws Exception {

    AppOffice365 appOffice365 = request.getContext().asType(AppOffice365.class);
    appOffice365 = Beans.get(AppOffice365Repository.class).find(appOffice365.getId());
    office365Service.syncContact(appOffice365);
    response.setView(
        ActionView.define("Office365 contact")
            .model(Partner.class.getName())
            .add("grid", "partner-grid")
            .add("form", "partner-form")
            .domain("self.office365Id IS NOT NULL")
            .map());
  }

  public void syncCalendar(ActionRequest request, ActionResponse response) throws Exception {

    AppOffice365 appOffice365 = request.getContext().asType(AppOffice365.class);
    appOffice365 = Beans.get(AppOffice365Repository.class).find(appOffice365.getId());
    office365Service.syncCalendar(appOffice365);
    response.setView(
        ActionView.define("Office365 contact")
            .model(ICalendar.class.getName())
            .add("grid", "calendar-grid")
            .add("form", "calendar-form")
            .domain("self.office365Id IS NOT NULL")
            .map());
  }

  public void syncMail(ActionRequest request, ActionResponse response) throws Exception {

    AppOffice365 appOffice365 = request.getContext().asType(AppOffice365.class);
    appOffice365 = Beans.get(AppOffice365Repository.class).find(appOffice365.getId());
    office365Service.syncMail(appOffice365);
    response.setView(
        ActionView.define("Office365 mail")
            .model(Message.class.getName())
            .add("grid", "message-grid")
            .add("form", "message-form")
            .domain("self.office365Id IS NOT NULL")
            .map());
  }
}
