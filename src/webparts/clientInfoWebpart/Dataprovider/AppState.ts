export class ClientInfoState {
  public ClientInformation = {
    LinkTitle: "",
    ClientNumber: "",
  };

  public UpdateState = (obj: ClientInfoState, data: any) => {
    obj.ClientInformation = data;
    return obj;
  }
}
