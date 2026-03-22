#ifndef _EXTERNHDL_H_
#define _EXTERNHDL_H_

#include <BaseExternHdl.hxx>

class ExternHdl : public BaseExternHdl
{
  public:
    ExternHdl(BaseExternHdl *nextHdl, PVSSulong funcCount, FunctionListRec fnList[])
      : BaseExternHdl(nextHdl, funcCount, fnList) {}

    const Variable *execute(ExecuteParamRec &param) override;
};

#endif
